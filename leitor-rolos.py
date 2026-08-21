import re
from datetime import datetime
import os
import win32com.client as win32
import tkinter as tk
from tkinter import ttk, messagebox

CAMINHO_EXCEL = r"Z:\Corte\Controle Consumo de Tecidos Duda.xlsm"
NOME_ABA = "AUDITORIA_ROLOS"

def limpar_texto_ascii(texto):
    # Remove caracteres invisíveis/de controle ASCII que causam erro no Excel
    return re.sub(r'[\x00-\x1f\x7f-\x9f]', '', texto)

def extrair_dados(texto_bruto):
    texto_limpo = limpar_texto_ascii(texto_bruto).replace('GS', '').strip()
    metros, peso, rolo = 0.0, 0.0, "N/I"
    
    # 1. Extrai Metragem (captura o decimal antes do prefixo 310)
    m_match = re.search(r'(?:311\d|211\d|12)(\d+\.\d+|\d{6})(?=310)', texto_limpo)
    if not m_match:
        m_match = re.search(r'(\d+\.\d+)(?=310)', texto_limpo)
        
    if m_match:
        val_str = m_match.group(1)
        metros = float(val_str) if '.' in val_str else float(val_str) / 100.0

    # 2. Extrai Peso (captura apenas o valor do peso até encontrar o prefixo '21' do rolo)
    p_match = re.search(r'310\d(\d+\.\d+|\d{1,6})(?=21)', texto_limpo)
    if not p_match:
        p_match = re.search(r'310\d(\d+\.\d+|\d+)', texto_limpo)

    if p_match:
        val_str = p_match.group(1)
        peso = float(val_str) if '.' in val_str else float(val_str) / 100.0

    # 3. Extrai Número do Rolo (14 dígitos após '21')
    r_match = re.search(r'21(\d{14})', texto_limpo)
    if r_match:
        rolo = r_match.group(1)
    else:
        r_fb = re.findall(r'\d{14}', texto_limpo)
        if r_fb:
            rolo = r_fb[-1]

    return rolo, metros, peso, texto_limpo

def salvar_no_excel(data_corte, rolo, metros, peso, texto_limpo):
    # Utiliza a aplicação nativa do Excel para preserver botões, formulários e macros VBA
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(CAMINHO_EXCEL)
    
    try:
        ws = wb.Worksheets(NOME_ABA)
    except Exception:
        ws = wb.Worksheets.Add()
        ws.Name = NOME_ABA

    # Encontra a primeira linha vazia com base no texto da Coluna A
    proxima_linha = 2
    while ws.Cells(proxima_linha, 1).Value is not None and str(ws.Cells(proxima_linha, 1).Value).strip() != "":
        proxima_linha += 1

    # Insere os dados nas colunas A, B, C, D e E
    ws.Cells(proxima_linha, 1).Value = data_corte
    ws.Cells(proxima_linha, 2).Value = rolo
    ws.Cells(proxima_linha, 3).Value = metros
    ws.Cells(proxima_linha, 4).Value = peso
    ws.Cells(proxima_linha, 5).Value = texto_limpo

    wb.Save()
    wb.Close(False)
    excel.Quit()

class AplicacaoLeitorPCP:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Consumo de Tecidos - PCP")
        self.root.geometry("650x550")
        self.root.configure(bg="#F4F6F9")

        self.total_bipados = 0

        # === TÍTULO ===
        lbl_titulo = tk.Label(root, text="LEITOR DE ROLOS - CORTE", font=("Helvetica", 16, "bold"), bg="#F4F6F9", fg="#1E293B")
        lbl_titulo.pack(pady=(15, 5))

        # === PAINEL SUPERIOR: DATA E CONTADOR ===
        frame_top = tk.Frame(root, bg="#FFFFFF", bd=1, relief="solid")
        frame_top.pack(fill="x", padx=20, pady=5)

        tk.Label(frame_top, text="📅 Data do Corte:", font=("Helvetica", 10, "bold"), bg="#FFFFFF").pack(side="left", padx=10, pady=10)
        
        self.ent_data = tk.Entry(frame_top, font=("Helvetica", 11), width=12, justify="center")
        self.ent_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.ent_data.pack(side="left", padx=5, pady=10)

        self.lbl_contador = tk.Label(frame_top, text="Rolos Gravados: 0", font=("Helvetica", 11, "bold"), bg="#FFFFFF", fg="#2563EB")
        self.lbl_contador.pack(side="right", padx=15, pady=10)

        # === ENTRADA DO BIP ===
        frame_bip = tk.Frame(root, bg="#F4F6F9")
        frame_bip.pack(fill="x", padx=20, pady=10)

        tk.Label(frame_bip, text="📌 Bipe a etiqueta aqui:", font=("Helvetica", 10, "bold"), bg="#F4F6F9").pack(anchor="w")
        
        self.ent_bip = tk.Entry(frame_bip, font=("Helvetica", 12), bd=2, relief="groove")
        self.ent_bip.pack(fill="x", pady=5)
        self.ent_bip.bind("<Return>", self.processar_bip)
        self.ent_bip.focus()

        # === MENSAGEM DE STATUS ===
        self.lbl_status = tk.Label(root, text="Aguardando leitura...", font=("Helvetica", 10, "italic"), bg="#F4F6F9", fg="#64748B")
        self.lbl_status.pack(pady=2)

        # === TABELA VISUAL (PREVIEW DOS ROLOS BIPADOS) ===
        frame_tabela = tk.Frame(root)
        frame_tabela.pack(fill="both", expand=True, padx=20, pady=10)

        colunas = ("rolo", "metros", "peso", "status")
        self.tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=8)
        
        self.tabela.heading("rolo", text="Nº Rolo")
        self.tabela.heading("metros", text="Metros (m)")
        self.tabela.heading("peso", text="Peso (kg)")
        self.tabela.heading("status", text="Status")

        self.tabela.column("rolo", width=180, anchor="center")
        self.tabela.column("metros", width=100, anchor="center")
        self.tabela.column("peso", width=100, anchor="center")
        self.tabela.column("status", width=180, anchor="center")

        self.tabela.pack(fill="both", expand=True)

        # === BOTÃO DE ENCERRAMENTO ===
        btn_concluir = tk.Button(root, text="🏁 Concluir e Fechar Sessão", font=("Helvetica", 11, "bold"), bg="#DC2626", fg="white", activebackground="#B91C1C", activeforeground="white", cursor="hand2", command=self.concluir_sessao)
        btn_concluir.pack(fill="x", padx=20, pady=(5, 15))

    def processar_bip(self, event=None):
        entrada = self.ent_bip.get().strip()
        self.ent_bip.delete(0, tk.END)

        if not entrada:
            return

        data_corte = self.ent_data.get().strip() or datetime.now().strftime("%d/%m/%Y")
        rolo, metros, peso, texto_limpo = extrair_dados(entrada)

        if not os.path.exists(CAMINHO_EXCEL):
            messagebox.showerror("Erro de Arquivo", f"Arquivo Excel não encontrado em:\n{CAMINHO_EXCEL}")
            return

        try:
            salvar_no_excel(data_corte, rolo, float(metros), float(peso), texto_limpo)

            self.total_bipados += 1
            self.lbl_contador.config(text=f"Rolos Gravados: {self.total_bipados}")
            self.lbl_status.config(text=f"✅ Rolo {rolo} gravado com sucesso!", fg="#16A34A")

            # Exibe na tabela da interface
            self.tabela.insert("", 0, values=(rolo, f"{metros:.2f} m", f"{peso:.2f} kg", "✅ Gravado"))

        except Exception as e:
            self.lbl_status.config(text="⚠️ ERRO: Feche o arquivo do Excel!", fg="#DC2626")
            messagebox.showwarning("Aviso", "O arquivo Excel está aberto por alguém!\nFeche o Excel e tente bipar novamente.")

    def concluir_sessao(self):
        messagebox.showinfo("Sessão Concluída", f"🏁 Sessão finalizada com sucesso!\n\n📅 Data do Corte: {self.ent_data.get().strip()}\n📦 Total de rolos gravados: {self.total_bipados}")
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = AplicacaoLeitorPCP(root)
    root.mainloop()
