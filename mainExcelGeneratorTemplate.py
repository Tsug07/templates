import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
import pdfplumber
import pandas as pd
import re
from collections import defaultdict

# Função para ler o Excel de contatos
def carregar_contatos_excel(caminho_excel):
    print()

# Função para extrair informações do PDF
def extrair_informacoes_pdf(caminho_pdf, contatos_dict):
   print()
            

# Função para gerar Excel a partir dos dados extraídos
def gerar_excel(dados, caminho_excel):
    df = pd.DataFrame(dados)
    df.to_excel(caminho_excel, index=False)
    print(f"Arquivo Excel criado: {caminho_excel}")

# Função para selecionar o arquivo PDF
def selecionar_pdf():
    caminho_pdf = filedialog.askopenfilename(
        title="Selecione o arquivo PDF",
        filetypes=(("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*"))
    )
    entrada_pdf.delete(0, tk.END)
    entrada_pdf.insert(0, caminho_pdf)
    
# Função para selecionar o caminho para salvar o Excel
def selecionar_destino_excel():
    caminho_excel = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")),
        title="Salvar arquivo Excel"
    )
    entrada_excel.delete(0, tk.END)
    entrada_excel.insert(0, caminho_excel)

# Função para selecionar o caminho para salvar o Excel
def selecionar_lista_contatos():
    caminho_contatos = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*"))
    )
    entrada_contatos.delete(0, tk.END)
    entrada_contatos.insert(0, caminho_contatos)

# Função para processar o PDF e gerar o Excel
def processar():

    caminho_pdf = entrada_pdf.get()
    caminho_excel = entrada_excel.get()
    caminho_contatos = entrada_contatos.get() #Excel Lista de contatos Onvio

    if not caminho_pdf or not caminho_excel or not caminho_contatos:
        messagebox.showwarning("Erro", "Por favor, selecione o arquivo PDF e o local para salvar o Excel.")
        return
    

def main():
    global entrada_pdf, entrada_excel, entrada_contatos 
     
    # Interface gráfica com Tkinter
    janela = tk.Tk()
    janela.title("Gerador de Excel Mensagem")
    janela.geometry("500x400")

    # Campo para o caminho do PDF
    lbl_pdf = tk.Label(janela, text="Selecione o arquivo PDF:")
    lbl_pdf.pack(pady=5)
    entrada_pdf = tk.Entry(janela, width=50)
    entrada_pdf.pack(pady=5)
    btn_pdf = tk.Button(janela, text="Selecionar PDF", command=selecionar_pdf)
    btn_pdf.pack(pady=5)

    # Campo para o caminho do Excel com a lista de contatos
    lbl_contatos = tk.Label(janela, text="Selecione o Excel de Contatos:")
    lbl_contatos.pack(pady=5)
    entrada_contatos = tk.Entry(janela, width=50)
    entrada_contatos.pack(pady=5)
    btn_contatos = tk.Button(janela, text="Selecionar Excel de Contatos", command=selecionar_lista_contatos)
    btn_contatos.pack(pady=5)

    # Campo para o caminho do Excel
    lbl_excel = tk.Label(janela, text="Selecione o destino do arquivo Excel:")
    lbl_excel.pack(pady=5)
    entrada_excel = tk.Entry(janela, width=50)
    entrada_excel.pack(pady=5)
    btn_excel = tk.Button(janela, text="Salvar como Excel", command=selecionar_destino_excel)
    btn_excel.pack(pady=5)

    # Botão para processar o PDF e gerar o Excel
    btn_processar = tk.Button(janela, text="Gerar Excel", command=processar)
    btn_processar.pack(pady=10)

    # Iniciar a janela
    janela.mainloop()

if __name__ == '__main__':
    main()