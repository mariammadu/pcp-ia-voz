import pyautogui
import time
import json
import base64
import urllib.request
import sounddevice as sd
import soundfile as sf
import numpy as np
from datetime import datetime

# ==========================================
# 1. CONFIGURAÇÃO DA CHAVE DO GEMINI
# ==========================================
API_KEY = "SUA_CHAVE_API"

# ==========================================
# 2. FUNÇÃO PARA GRAVAR O MICROFONE (INTERATIVA)
# ==========================================
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
        # A gravação roda enquanto esta caixa de mensagem estiver aberta na tela
        pyautogui.alert(
            text="🎙️ GRAVANDO VOZ...\n\nFale o seu lançamento e, assim que terminar de falar, clique em 'OK' (ou aperte Enter) para processar.",
            title="PCP Voz - Gravação"
        )
        gravando = False # Para a captura imediatamente ao clicar em OK

    # Concatena os fragmentos de áudio e salva no arquivo .wav
    audio_completo = np.concatenate(dados_audio, axis=0)
    sf.write(arquivo, audio_completo, samplerate)
    
    print("[✅] Gravação concluída com sucesso!")
    return arquivo

# ==========================================
# 3. FUNÇÃO PARA PREENCHER O EXCEL
# ==========================================
def preencher_excel(dados):
    print("\n--- INICIANDO PREENCHIMENTO NO EXCEL ---")
    
    # 1. Desproteger planilha (Alt + R + P)
    pyautogui.hotkey('alt', 'r', 'p')
    time.sleep(0.5)

    # 2. Preencher Topo (Data -> Modelo -> Medida -> Camada)
    pyautogui.write(str(dados.get("data", "")), interval=0.03)
    pyautogui.press("tab")

    pyautogui.write(str(dados.get("modelo", "")), interval=0.03)
    pyautogui.press("tab")

    pyautogui.write(str(dados.get("medida", "")), interval=0.03)
    pyautogui.press("tab")

    pyautogui.write(str(dados.get("camada", "")), interval=0.03)
    pyautogui.press("tab")

    # 3. Preencher Grade de Tamanhos
    tamanhos = ["PP", "P", "M", "G", "GG", "EX", "EXG", "2G", "3G", "4G", "5G", "Unico"]

    for tamanho in tamanhos:
        valor = dados.get(tamanho, "")
        if valor != "":
            pyautogui.write(str(valor), interval=0.03)
        pyautogui.press("tab")

    time.sleep(0.5)

    # 4. Salvar Lançamento no formulário (Enter)
    pyautogui.press("enter")
    time.sleep(1)

    # 5. Reproteger planilha (Alt + R + P)
    pyautogui.hotkey('alt', 'r', 'p')
    print("--- CONCLUÍDO COM SUCESSO! ---")

# ==========================================
# 4. EXECUÇÃO PRINCIPAL
# ==========================================
if __name__ == "__main__":
    # 1. Grava o áudio no tempo exato em que o usuário fala
    arquivo_audio = gravar_audio()
    
    print("Enviando áudio para o Gemini processar...")
    
    # Converte o arquivo de áudio gravado para Base64
    with open(arquivo_audio, "rb") as f:
        audio_b64 = base64.b64encode(f.read()).decode('utf-8')

    # Obter referência da data atual dinamicamente
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

    # Endpoint oficial do Gemini
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent"
    headers = {
        "Content-Type": "application/json",
        "X-goog-api-key": API_KEY
    }

    payload = {
        "contents": [
            {
                "parts": [
                    {
                        "inline_data": {
                            "mime_type": "audio/wav",
                            "data": audio_b64
                        }
                    },
                    {
                        "text": prompt
                    }
                ]
            }
        ]
    }

    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers=headers, method='POST')

    try:
        # Timeout de 30s evita travamentos por falhas de rede
        with urllib.request.urlopen(req, timeout=30) as response:
            res = json.loads(response.read().decode('utf-8'))
            texto_resposta = res['candidates'][0]['content']['parts'][0]['text']
            
            # Limpa formatações Markdown
            texto_limpo = texto_resposta.replace("```json", "").replace("```", "").strip()
            dados_extraidos = json.loads(texto_limpo)
            
            print("\nO Gemini ouviu e entendeu estes dados:")
            print(json.dumps(dados_extraidos, indent=2, ensure_ascii=False))
            
            # Trava de segurança para garantir o foco no Excel antes de digitar
            pyautogui.alert(
                text="Os dados foram processados com sucesso!\n\n1. Clique no campo DATA do formulário no seu Excel.\n2. Em seguida, clique em OK nesta janela.",
                title="🤖 Automação Pronta"
            )
            time.sleep(1.5)  # Tempo para o Windows focar no Excel
            
            preencher_excel(dados_extraidos)

    except Exception as e:
        print("\n❌ Erro ao processar áudio:", e)
