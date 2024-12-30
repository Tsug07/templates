import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import threading
import time
import os
import psutil
from datetime import datetime
import re
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException

"""
Automação de mensagem padrão de cobrança no Onvio Messenger para clientes que estão com pagamento atrasado 

"""

# Variável global para controlar o cancelamento do processamento
cancelar = False

#FUNÇÕES DO PROGRAMA
def focar_barra_endereco_e_navegar(driver, termo_busca):
    atualizar_log()

def encerrar_processos_chrome():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] == 'chrome.exe':
            try:
                proc.terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    time.sleep(2)  # Aguarda um pouco para garantir que os processos foram encerrados

def abrir_chrome_com_url(url):
    encerrar_processos_chrome()
    
    # Caminho para o perfil padrão do Chrome
    user_data_dir = os.path.expanduser('~') + r'\AppData\Local\Google\Chrome\User Data'
    
    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument(f"user-data-dir={user_data_dir}")
    chrome_options.add_argument("--profile-directory=Default")
    chrome_options.add_argument("--disable-translate")  # Tenta desabilitar a tradução automática
    chrome_options.add_argument("--lang=pt-BR")  # Define o idioma para português do Brasil
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_experimental_option("prefs", {
        "translate": {"enabled": "false"},
        "profile.default_content_setting_values.notifications": 2
    })
    
    
    service = Service(ChromeDriverManager().install())
    
    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
        driver.set_page_load_timeout(180)  # Aumenta o tempo limite para carregar a página
        driver.get(url)
        atualizar_log(f"Chrome aberto com a URL: {url}")
        return driver
    except Exception as e:
        atualizar_log(f"Erro ao abrir o Chrome: {str(e)}")
        return None

def esperar_carregamento_completo(driver):
    global cancelar
    cancelar = False


def processar_resultados_busca(driver):
    global cancelar
    cancelar = False


def focar_barra_mensagem_enviar(driver, mensagem):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return


def encontrar_e_clicar_barra_contatos(driver, contato, grupo):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
 
def clicar_voltar_lista_contatos(driver):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return

    
def focar_pagina(driver):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    


    
def mensagemPadrao(valores, datas, nome, carta):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return
    atualizar_log('')


def ler_dados_excel(caminho_excel):
    atualizar_log('')

def extrair_cod_nome_contatos_e_grupos(dados):
   atualizar_log('')


        
# Função para selecionar o arquivo Excel
def selecionar_excel():
    arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo Excel",
        filetypes=(("Arquivos Excel", "*.xlsx *.xls"), ("Todos os arquivos", "*.*"))
    )
    if arquivo:
        caminho_excel.set(arquivo)
        atualizar_log(f"Arquivo Excel selecionado: {arquivo}")

# Função para iniciar o processamento dos dados
def iniciar_processamento():
    global cancelar
    cancelar = False
    
    excel = caminho_excel.get()
    if excel:
        atualizar_log("Iniciando processamento...")
        botao_iniciar.config(state=tk.DISABLED)  # Desabilitar o botão para evitar múltiplos cliques
        thread = threading.Thread(target=processar_dados, args=(excel,))
        thread.start()
    else:
        messagebox.showwarning("Atenção", "Por favor, selecione o arquivo Excel.")

# Função do seu código que já está pronta para processar os dados
def processar_dados(excel):
    global cancelar
    if cancelar:
        atualizar_log("Processamento cancelado!", cor="azul")
        return 
    
    # Arquivo Excel
    caminho_excel = excel
    
    
    
# Função para cancelar o processamento
def cancelar_processamento():
    global cancelar
    cancelar = True
    atualizar_log("Cancelando processamento...")
    botao_fechar.config(state=tk.NORMAL)  # Habilitar o botão de fechar o programa

# Função para cancelar e fechar o programa
def fechar_programa():
    janela.quit()

# Função para finalizar o programa com uma mensagem
def finalizar_programa():
    messagebox.showinfo("Processo Finalizado", "O processamento foi concluído com sucesso!")
    botao_fechar.config(state=tk.NORMAL)  # Habilitar o botão de fechar o programa
    botao_iniciar.config(state=tk.NORMAL)  # Reabilitar o botão de iniciar

# Função para atualizar o log na área de texto
def atualizar_log(mensagem, cor=None):
    log_text.config(state=tk.NORMAL)  # Habilitar edição temporária
    if cor == "vermelho":
        log_text.insert(tk.END, mensagem + "\n", "vermelho")  # Inserir nova mensagem com tag 'vermelho'
    elif cor == "verde":
        log_text.insert(tk.END, mensagem + "\n", "verde")  # Inserir nova mensagem com tag 'verde'
    elif cor == "azul":
        log_text.insert(tk.END, mensagem + "\n", "azul")
    else:
        log_text.insert(tk.END, mensagem + "\n")  # Inserir nova mensagem sem tag
    log_text.config(state=tk.DISABLED)  # Desabilitar edição novamente
    log_text.see(tk.END)  # Scroll automático para a última linha

# Função para configurar a tag de cor no log
def configurar_tags_log():
    log_text.tag_config("vermelho", foreground="red")  # Configura a cor vermelha para a tag 'vermelho'
    log_text.tag_config("verde", foreground="green")  # Configura a cor vermelha para a tag 'verde'
    log_text.tag_config("azul", foreground="blue") # Configura a cor azul para a tag 'azul'
# Função main para encapsular a lógica do programa
def main():
    global janela, caminho_pasta, caminho_excel, botao_fechar, botao_iniciar, log_text

    

    # Criar a janela principal
    janela = tk.Tk()
    janela.title("Envio Mensagem Onvio")
    janela.geometry("600x400")
    janela.resizable(False, False)

    # Variáveis para armazenar os caminhos
    caminho_pasta = tk.StringVar()
    caminho_excel = tk.StringVar()

    # Frame para seleção de pasta e arquivo
    frame_selecao = tk.Frame(janela)
    frame_selecao.pack(pady=10)


    # Label e Botão para selecionar o arquivo Excel
    label_excel = tk.Label(frame_selecao, text="Arquivo Excel:")
    label_excel.grid(row=1, column=0, pady=5, padx=5)

    entrada_excel = tk.Entry(frame_selecao, textvariable=caminho_excel, width=50, state='readonly')
    entrada_excel.grid(row=1, column=1, padx=5)

    botao_excel = tk.Button(frame_selecao, text="Selecionar Excel", command=selecionar_excel)
    botao_excel.grid(row=1, column=2, padx=5)

    # Botão para iniciar o processamento
    botao_iniciar = tk.Button(janela, text="Iniciar Processamento", command=iniciar_processamento)
    botao_iniciar.pack(pady=10)

    # Botão para cancelar e fechar o programa
    botao_cancelar = tk.Button(janela, text="Cancelar Processamento", command=cancelar_processamento)
    botao_cancelar.pack(pady=5)

    # Botão para fechar o programa (desabilitado até o processamento terminar)
    botao_fechar = tk.Button(janela, text="Fechar Programa", command=fechar_programa, state=tk.DISABLED)
    botao_fechar.pack(pady=5)

    # Frame para o log
    frame_log = tk.Frame(janela)
    frame_log.pack(pady=10, fill=tk.BOTH, expand=True)

    # Área de texto para o log com barra de rolagem
    log_text = scrolledtext.ScrolledText(frame_log, wrap=tk.WORD, height=10, state=tk.DISABLED)
    log_text.pack(fill=tk.BOTH, expand=True)

    # Configurar as tags de cor para o log
    configurar_tags_log()
    
    # Iniciar o loop da interface
    janela.mainloop()

# Garantir que o código só execute se este arquivo for o principal
if __name__ == '__main__':
    main()
