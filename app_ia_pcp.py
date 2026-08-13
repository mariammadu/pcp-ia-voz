import pyautogui
import time
import json
import base64
import urllib.request
import sounddevice as sd
import soundfile as sf

# ==========================================
# 1. CONFIGURAÇÃO DA CHAVE DO GEMINI
# ==========================================
API_KEY = "SUA_CHAVE_API_AQUI"

# ==========================================
# 2. FUNÇÃO PARA GRAVAR O MICROFONE
# ==========================================
def gravar_audio(duracao=25, arquivo="gravacao.wav", samplerate=44100):
    print("\n[🎙️] LIGANDO MICROFONE... FALE AGORA!")
    pyautogui.alert(
        text=f"Assim que clicar em OK, FALE o seu lançamento no microfone!\nVocê terá {duracao} segundos para falar.",
        title="🎙️ Gravando Voz..."
    )
    
    audio = sd.rec(int(duracao * samplerate), samplerate=samplerate, channels=1, dtype='int16')
    sd.wait()
    sf.write(arquivo, audio, samplerate)
    print("[✅] Gravação concluída!")
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
    # 1. Grava o áudio
    arquivo_audio = gravar_audio(duracao=25)
    
    print("Enviando áudio para o Gemini processar...")
    
    # Converte o arquivo de áudio gravado para Base64
    with open(arquivo_audio, "rb") as f:
        audio_b64 = base64.b64encode(f.read()).decode('utf-8')

    prompt = """
    Ouça este áudio de PCP têxtil e extraia os dados para o formulário.
    Retorne APENAS um JSON com esta estrutura exata (sem markdown, sem texto extra):
    {
      "data": "DD/MM/AAAA",
      "modelo": "número",
      "medida": "medida em metros",
      "camada": "quantidade de camadas",
      "PP": "", "P": "", "M": "", "G": "", "GG": "",
      "EX": "", "EXG": "", "2G": "", "3G": "", "4G": "", "5G": "", "Unico": ""
    }
    Se a pessoa não falar a data no áudio, insira a data atual.
    """

    # Endpoint oficial validado
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
        with urllib.request.urlopen(req) as response:
            res = json.loads(response.read().decode('utf-8'))
            texto_resposta = res['candidates'][0]['content']['parts'][0]['text']
            
            # Limpa qualquer tag de código Markdown que a IA colocar
            texto_limpo = texto_resposta.replace("```json", "").replace("```", "").strip()
            dados_extraidos = json.loads(texto_limpo)
            
            print("\nO Gemini ouviu e entendeu estes dados:")
            print(json.dumps(dados_extraidos, indent=2, ensure_ascii=False))
            
            pyautogui.alert("A IA processou sua voz! Clique no campo DATA do formulário no Excel e clique em OK aqui.")
            time.sleep(1)
            
            preencher_excel(dados_extraidos)

    except Exception as e:
        print("\n❌ Erro ao processar áudio:", e)