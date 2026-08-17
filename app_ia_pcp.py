import pyautogui
import time
import json
import sounddevice as sd
import soundfile as sf
import numpy as np
from datetime import datetime
import traceback
from google import genai
from google.genai import types


API_KEY = "SUA_CHAVE_API"

client = genai.Client(api_key=API_KEY)

def gravar_audio(arquivo="gravacao.wav", samplerate=44100):
    print("\n[🎙️] LIGANDO MICROFONE...")
    
    gravando = True
    dados_audio = []

    def callback(indata, frames, time_info, status):
        if gravando:
            dados_audio.append(indata.copy())

    # Inicia a captura de áudio em segundo plano
    stream = sd.InputStream(samplerate=samplerate, channels=1, dtype='int16', callback=callback)
    with stream:
        pyautogui.alert(
            text="🎙️ GRAVANDO VOZ...\n\nFale o seu lançamento e, assim que terminar de falar, clique em 'OK' (ou aperte Enter) para processar.",
            title="PCP Voz - Gravação"
        )
        gravando = False

    audio_completo = np.concatenate(dados_audio, axis=0)
    sf.write(arquivo, audio_completo, samplerate)
    
    print("[✅] Gravação concluída com sucesso!")
    return arquivo

def preencher_excel(dados):
    print("\n--- INICIANDO PREENCHIMENTO NO EXCEL ---")
    
   
    pyautogui.hotkey('alt', 'r', 'p')
    time.sleep(0.5)

    
    pyautogui.write(str(dados.get("data", "")), interval=0.03)
    pyautogui.press("tab")

    pyautogui.write(str(dados.get("modelo", "")), interval=0.03)
    pyautogui.press("tab")

    pyautogui.write(str(dados.get("medida", "")), interval=0.03)
    pyautogui.press("tab")

    pyautogui.write(str(dados.get("camada", "")), interval=0.03)
    pyautogui.press("tab")

   
    tamanhos = ["PP", "P", "M", "G", "GG", "EX", "EXG", "2G", "3G", "4G", "5G", "Unico"]

    for tamanho in tamanhos:
        valor = dados.get(tamanho, "")
        if valor != "":
            pyautogui.write(str(valor), interval=0.03)
        pyautogui.press("tab")

    time.sleep(0.5)

    
    pyautogui.press("enter")
    time.sleep(1)

    
    pyautogui.hotkey('alt', 'r', 'p')
    print("--- CONCLUÍDO COM SUCESSO! ---")


if __name__ == "__main__":
    try:
        arquivo_audio = gravar_audio()
        
        print("Enviando áudio para o Gemini processar via SDK oficial...")
        
        
        with open(arquivo_audio, "rb") as f:
            audio_bytes = f.read()

        hoje = datetime.now()
        data_hoje_str = hoje.strftime("%d/%m/%Y")
        ano_atual = hoje.year

        prompt = f"""
        Ouça este áudio de PCP têxtil e extraia os dados para o formulário.
        Retorne APENAS um JSON com esta estrutura exata (sem markdown, sem texto extra):
        {{
          "data": "DD/MM/AAAA",
          "modelo": "número",
          "medida": "medida em metros",
          "camada": "quantidade de camadas",
          "PP": "", "P": "", "M": "", "G": "", "GG": "",
          "EX": "", "EXG": "", "2G": "", "3G": "", "4G": "", "5G": "", "Unico": ""
        }}

        REGRAS RÍGIDAS PARA O CAMPO DATA:
        1. A data atual de referência de hoje é: {data_hoje_str} (Ano Atual: {ano_atual}).
        2. SE O USUÁRIO FALAR O ANO (ex: "2024", "2025", "2026"): Extraia EXATAMENTE o ano falado no áudio.
        3. SE O USUÁRIO OMITIR O ANO (ex: disse apenas "12 de maio"): Utilize o ano atual da referência ({ano_atual}).
        4. SE O USUÁRIO NÃO MENCIONAR NENHUMA DATA: Insira a data atual completa ({data_hoje_str}).
        """

        
        response = client.models.generate_content(
        model="gemini-3.6-flash",
        contents=[
        types.Part.from_bytes(
            data=audio_bytes,
            mime_type="audio/wav"
        ),
        prompt
    ]
)

        texto_resposta = response.text
        texto_limpo = texto_resposta.replace("```json", "").replace("```", "").strip()
        dados_extraidos = json.loads(texto_limpo)
        
        print("\nO Gemini ouviu e entendeu estes dados:")
        print(json.dumps(dados_extraidos, indent=2, ensure_ascii=False))
        
        pyautogui.alert(
            text="Os dados foram processados com sucesso!\n\n1. Clique no campo DATA do formulário no seu Excel.\n2. Em seguida, clique em OK nesta janela.",
            title="🤖 Automação Pronta"
        )
        time.sleep(1.5)
        
        preencher_excel(dados_extraidos)

    except Exception as e:
        erro_detalhado = traceback.format_exc()
        print("\n ERRO DETECTADO:\n", erro_detalhado)
        pyautogui.alert(
            text=f"Ocorreu um erro ao executar a automação:\n\n{e}\n\nDetalhes:\n{erro_detalhado[:300]}...",
            title=" Erro na Execução"
        )
