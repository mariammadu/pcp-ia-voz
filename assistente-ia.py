import os
import re
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from google import genai

# ==========================================
# CONFIGURAÇÕES DE CAMINHO E CHAVE API
# ==========================================
CAMINHO_EXCEL = r"Z:\Corte\Controle Consumo de Tecidos Duda.xlsm"

# Insira aqui a sua chave de API do Google Gemini
API_KEY = "CHAVEAPI" 

client = genai.Client(api_key=API_KEY)

def ler_dados_excel():
    """Lê todas as abas da planilha local e prepara a estrutura para a IA."""
    if not os.path.exists(CAMINHO_EXCEL):
        return None, f"Arquivo não encontrado no caminho:\n{CAMINHO_EXCEL}"
    
    try:
        xls = pd.ExcelFile(CAMINHO_EXCEL)
        contexto = []

        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            
            # Limpa linhas e colunas vazias
            df = df.dropna(how='all').dropna(how='all', axis=1)
            
            # Limita a leitura das últimas 250 linhas para otimizar envio
            if len(df) > 250:
                df = df.tail(250)

            contexto.append(f"=== ABA: {sheet_name} ===")
            contexto.append(df.to_string(index=False))
            contexto.append("\n")

        return "\n".join(contexto), None
    except Exception as e:
        return None, f"Erro ao acessar a planilha: {str(e)}"

def perguntar_ia(pergunta, contexto_planilha):
    """Envia o contexto formatado da planilha e a dúvida para o modelo Gemini."""
    prompt = f"""
    Você é o Assistente Virtual especialista em PCP (Planejamento e Controle de Produção) da fábrica têxtil.
    Abaixo estão os dados reais extraídos diretamente das abas da planilha do Excel:

    {contexto_planilha}

    PERGUNTA DO USUÁRIO: "{pergunta}"

    DIRETRIZES DE RESPOSTA:
    1. Responda baseando-se EXCLUSIVAMENTE nos dados fornecidos na planilha acima.
    2. Utilize termos do setor têxtil (metros, peso/kg, rolos, modelos, camadas, peças).
    3. Seja direto, conciso e forneça valores exatos sempre que solicitado (ex: totais somados, médias ou contagens).
    4. Se a informação não existir na planilha, informe com clareza que o dado não consta nos registros.
    """

    modelos = ["gemini-2.5-flash", "gemini-1.5-flash"]
    
    for modelo in modelos:
        try:
            response = client.models.generate_content(
                model=modelo,
                contents=prompt
            )
            return response.text
        except Exception:
            continue

    return "❌ Falha na conexão com a IA. Verifique sua chave API ou conexão de internet."

# ==========================================
# INTERFACE GRÁFICA TKINTER
# ==========================================
class AppConsultaIA:
    def __init__(self, root):
        self.root = root
        self.root.title("Assistente Virtual PCP - Consulta por IA")
        self.root.geometry("750x600")
        self.root.configure(bg="#F4F6F9")

        # Título
        lbl_titulo = tk.Label(root, text="🤖 CONSULTA IA - PCP TÊXTIL", font=("Helvetica", 16, "bold"), bg="#F4F6F9", fg="#1E293B")
        lbl_titulo.pack(pady=(15, 2))

        lbl_sub = tk.Label(root, text="Faça perguntas sobre lançamentos, auditoria de rolos, modelos e consumos.", font=("Helvetica", 9), bg="#F4F6F9", fg="#64748B")
        lbl_sub.pack(pady=(0, 10))

        # Entrada da Pergunta
        frame_input = tk.Frame(root, bg="#F4F6F9")
        frame_input.pack(fill="x", padx=20, pady=5)

        tk.Label(frame_input, text="Digite sua dúvida:", font=("Helvetica", 10, "bold"), bg="#F4F6F9").pack(anchor="w")
        
        self.ent_pergunta = tk.Entry(frame_input, font=("Helvetica", 11), bd=2, relief="groove")
        self.ent_pergunta.pack(fill="x", pady=5)
        self.ent_pergunta.bind("<Return>", self.executar_consulta)

        self.btn_consultar = tk.Button(frame_input, text="🔍 Consultar IA", font=("Helvetica", 10, "bold"), bg="#2563EB", fg="white", activebackground="#1D4ED8", activeforeground="white", cursor="hand2", command=self.executar_consulta)
        self.btn_consultar.pack(anchor="e", pady=5)

        # Caixa de Resposta
        frame_resposta = tk.Frame(root, bg="#F4F6F9")
        frame_resposta.pack(fill="both", expand=True, padx=20, pady=10)

        tk.Label(frame_resposta, text="Resposta:", font=("Helvetica", 10, "bold"), bg="#F4F6F9").pack(anchor="w")

        self.txt_resposta = tk.Text(frame_resposta, font=("Helvetica", 10), wrap="word", bd=2, relief="groove", bg="#FFFFFF")
        self.txt_resposta.pack(fill="both", expand=True, pady=5)

    def executar_consulta(self, event=None):
        pergunta = self.ent_pergunta.get().strip()
        if not pergunta:
            return

        self.txt_resposta.delete("1.0", tk.END)
        self.txt_resposta.insert(tk.END, "⏳ Lendo dados da planilha e processando sua consulta via IA...\n")
        self.root.update()

        contexto, erro = ler_dados_excel()
        if erro:
            self.txt_resposta.delete("1.0", tk.END)
            self.txt_resposta.insert(tk.END, f"❌ ERRO: {erro}")
            return

        resposta = perguntar_ia(pergunta, contexto)

        self.txt_resposta.delete("1.0", tk.END)
        self.txt_resposta.insert(tk.END, resposta)

if __name__ == "__main__":
    root = tk.Tk()
    app = AppConsultaIA(root)
    root.mainloop()
