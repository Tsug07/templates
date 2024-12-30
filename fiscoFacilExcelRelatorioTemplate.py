import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
from datetime import datetime

def selecionar_pasta_pdf():
    """Abre o seletor de pasta para escolher a pasta contendo os PDFs."""
    pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs")
    if pasta:
        entrada_pasta_pdf.set(pasta)
        adicionar_log(f"Pasta selecionada: {pasta}")

def selecionar_arquivo_excel():
    """Abre o seletor de arquivo para escolher onde salvar o Excel."""
    arquivo = filedialog.asksaveasfilename(
        title="Salvar como",
        defaultextension=".xlsx",
        filetypes=[("Arquivo Excel", "*.xlsx")],
    )
    if arquivo:
        entrada_arquivo_excel.set(arquivo)
        adicionar_log(f"Arquivo Excel selecionado: {arquivo}")

def gerar_excel():
    """Função chamada ao clicar no botão 'Gerar Excel'."""
    pasta_pdf = entrada_pasta_pdf.get()
    caminho_excel = entrada_arquivo_excel.get()

    if not pasta_pdf:
        messagebox.showerror("Erro", "Selecione a pasta com os PDFs.")
        adicionar_log("Erro: Nenhuma pasta selecionada.")
        return

    if not caminho_excel:
        messagebox.showerror("Erro", "Selecione onde salvar o arquivo Excel.")
        adicionar_log("Erro: Nenhum local para salvar o Excel selecionado.")
        return

    # Simulação do processamento
    try:
        adicionar_log("Iniciando processamento dos PDFs...")
        # Substituir com sua função de processamento
        # processar_pdfs_e_salvar_excel(pasta_pdf, caminho_excel)
        adicionar_log(f"Excel gerado com sucesso em: {caminho_excel}")
        messagebox.showinfo("Sucesso", f"Excel gerado com sucesso em:\n{caminho_excel}")
    except Exception as e:
        mensagem_erro = f"Erro ao gerar o Excel: {str(e)}"
        messagebox.showerror("Erro", mensagem_erro)
        adicionar_log(mensagem_erro)

def fechar_programa():
    """Fecha o programa."""
    janela.destroy()

def adicionar_log(mensagem):
    """Adiciona uma mensagem ao log."""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_mensagem = f"[{timestamp}] {mensagem}\n"
    log_text.insert(tk.END, log_mensagem)
    log_text.see(tk.END)  # Scroll automático para a última mensagem

def salvar_log():
    """Salva o log em um arquivo de texto."""
    arquivo = filedialog.asksaveasfilename(
        title="Salvar log como",
        defaultextension=".txt",
        filetypes=[("Arquivo de texto", "*.txt")],
    )
    if arquivo:
        with open(arquivo, "w", encoding="utf-8") as f:
            f.write(log_text.get("1.0", tk.END))
        adicionar_log(f"Log salvo em: {arquivo}")
        messagebox.showinfo("Sucesso", f"Log salvo em:\n{arquivo}")

# Criação da janela principal
janela = tk.Tk()
janela.title("Gerar Excel de PDFs")
janela.geometry("600x500")

# Variáveis de entrada
entrada_pasta_pdf = tk.StringVar()
entrada_arquivo_excel = tk.StringVar()

# Labels e campos de entrada
tk.Label(janela, text="Selecione a pasta com os PDFs:").pack(pady=5)
tk.Entry(janela, textvariable=entrada_pasta_pdf, width=60).pack(pady=5)
tk.Button(janela, text="Selecionar Pasta", command=selecionar_pasta_pdf).pack(pady=5)

tk.Label(janela, text="Selecione onde salvar o Excel:").pack(pady=5)
tk.Entry(janela, textvariable=entrada_arquivo_excel, width=60).pack(pady=5)
tk.Button(janela, text="Selecionar Arquivo", command=selecionar_arquivo_excel).pack(pady=5)

# Botões de ação
tk.Button(janela, text="Gerar Excel", command=gerar_excel, bg="green", fg="white").pack(pady=10)
tk.Button(janela, text="Fechar", command=fechar_programa, bg="red", fg="white").pack(pady=5)

# Log
tk.Label(janela, text="Log do processo:").pack(pady=5)
log_text = scrolledtext.ScrolledText(janela, width=70, height=15, state=tk.NORMAL)
log_text.pack(pady=5)

# Botão para salvar log
tk.Button(janela, text="Salvar Log", command=salvar_log).pack(pady=5)

# Loop principal da interface
janela.mainloop()
