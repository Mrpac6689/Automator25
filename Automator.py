import os
import csv
import subprocess
import platform
import unicodedata
import time
import pandas as pd
import re
import configparser
import pygetwindow as gw
import ctypes
import tkinter as tk
import threading
import sys
import requests
import zipfile
import shutil
import xml.etree.ElementTree as ET
from tkinter import messagebox, filedialog, scrolledtext
from tkinter.scrolledtext import ScrolledText
from colorama import init, Fore, Style
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path

# Para captura da tecla espaço
if platform.system() == "Windows":
    import msvcrt
else:
    import sys, tty, termios

# Inicializa a colorama para cores no terminal
init(autoreset=True)

# Função para limpar a tela
def limpar_tela():
    sistema_operacional = platform.system()
    if sistema_operacional == "Windows":
        os.system("cls")
    else:
        os.system("clear")

# Função para exibir o cabeçalho do programa
def exibir_cabecalho():
    print(Fore.CYAN + "| * -------------------------------------------------------------*|")
    print(Fore.CYAN + "|                        AUTOMATOR - AutoReg                      |")
    print(Fore.CYAN + "|            Automação de pacientes a dar alta - SISREG           |")
    print(Fore.CYAN + "|                  Versão 2.0 - Setembro de 2024                  |")
    print(Fore.CYAN + "|                Michel R. Paes - Turbinado por ChatGPT           |")
    print(Fore.CYAN + "|                       michelrpaes@gmail.com                     |")
    print(Fore.CYAN + "| * -------------------------------------------------------------*|")

# Função para exibir o menu de opções
def exibir_menu():
    print(Fore.YELLOW + "\n[1] Extrair internados SISREG")
    print(Fore.YELLOW + "[2] Extrair internados G-HOSP")
    print(Fore.YELLOW + "[3] Comparar e tratar dados")
    print(Fore.YELLOW + "[4] Capturar motivo da alta")
    print(Fore.YELLOW + "[5] Sair")

# Função para normalizar o nome (remover acentos, transformar em minúsculas)
def normalizar_nome(nome):
    # Remove acentos e transforma em minúsculas
    nfkd = unicodedata.normalize('NFKD', nome)
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower()

# Função para esperar pela tecla espaço
def esperar_tecla_espaco():
    print(Fore.CYAN + "\nPressione espaço para continuar...")
    
    if platform.system() == "Windows":
        while True:
            if msvcrt.getch() == b' ':
                break
    else:
        fd = sys.stdin.fileno()
        old_settings = termios.tcgetattr(fd)
        try:
            tty.setraw(sys.stdin.fileno())
            while True:
                ch = sys.stdin.read(1)
                if ch == ' ':
                    break
        finally:
            termios.tcsetattr(fd, termios.TCSADRAIN, old_settings)

# Função para comparar os arquivos CSV e salvar os pacientes a dar alta
def comparar_dados():
    # Caminho para os arquivos
    arquivo_sisreg = 'internados_sisreg.csv'
    arquivo_ghosp = 'internados_ghosp.csv'

    # Verifica se os arquivos existem
    if not os.path.exists(arquivo_sisreg) or not os.path.exists(arquivo_ghosp):
        print(Fore.RED + "\nOs arquivos internados_sisreg.csv ou internados_ghosp.csv não foram encontrados!")
        return

    # Lê os arquivos CSV
    with open(arquivo_sisreg, 'r', encoding='utf-8') as sisreg_file:
        sisreg_nomes_lista = [normalizar_nome(linha[0].strip()) for linha in csv.reader(sisreg_file) if linha]

    # Ignora a primeira linha (cabeçalho)
    sisreg_nomes = set(sisreg_nomes_lista[1:])

    with open(arquivo_ghosp, 'r', encoding='utf-8') as ghosp_file:
        ghosp_nomes = {normalizar_nome(linha[0].strip()) for linha in csv.reader(ghosp_file) if linha}

    # Encontra os pacientes a dar alta (presentes no SISREG e ausentes no G-HOSP)
    pacientes_a_dar_alta = sisreg_nomes - ghosp_nomes

    if pacientes_a_dar_alta:
        print(Fore.GREEN + "\n---===> PACIENTES A DAR ALTA <===---")
        for nome in sorted(pacientes_a_dar_alta):
            print(Fore.LIGHTYELLOW_EX + nome)  # Alterado para amarelo neon
        
        # Salva a lista em um arquivo CSV
        with open('pacientes_de_alta.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Nome'])  # Cabeçalho
            for nome in sorted(pacientes_a_dar_alta):
                writer.writerow([nome])
        
        print(Fore.CYAN + "\nA lista de pacientes a dar alta foi salva em 'pacientes_de_alta.csv'.")
        esperar_tecla_espaco()
    else:
        print(Fore.RED + "\nNenhum paciente a dar alta encontrado!")
        esperar_tecla_espaco()

### Definições extrator.py

# Função para ler as credenciais do arquivo config.ini
def ler_credenciais():
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    usuario_sisreg = config['SISREG']['usuario']
    senha_sisreg = config['SISREG']['senha']
    
    return usuario_sisreg, senha_sisreg

def extrator():
    # Exemplo de uso no script extrator.py
    usuario, senha = ler_credenciais()

    # Caminho para o ChromeDriver
    chrome_driver_path = "chromedriver.exe"

    # Cria um serviço para o ChromeDriver
    service = Service(executable_path=chrome_driver_path)

    # Modo silencioso
    chrome_options = Options()
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--incognito')

    # Inicializa o navegador (Chrome neste caso) usando o serviço
    driver = webdriver.Chrome(service=service, options=chrome_options)

    # Minimizando a janela após iniciar o Chrome
    driver.minimize_window()

    # Acesse a página principal do SISREG
    driver.get('https://sisregiii.saude.gov.br/')

    try:
        print("Tentando localizar o campo de usuário...")
        usuario_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "usuario"))
        )
        print("Campo de usuário encontrado!")
    
        print("Tentando localizar o campo de senha...")
        senha_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "senha"))
        )
        print("Campo de senha encontrado!")

        # Preencha os campos de login
        print("Preenchendo o campo de usuário...")
        usuario_field.send_keys(usuario)
        
        print("Preenchendo o campo de senha...")
        senha_field.send_keys(senha)

        # Pressiona o botão de login
        print("Tentando localizar o botão de login...")
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='entrar' and @value='entrar']"))
        )
        
        print("Botão de login encontrado. Tentando fazer login...")
        login_button.click()

        time.sleep(5)
        print("Login realizado com sucesso!")

        # Agora, clica no link "Saída/Permanência"
        print("Tentando localizar o link 'Saída/Permanência'...")
        saida_permanencia_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[@href='/cgi-bin/config_saida_permanencia' and text()='saída/permanência']"))
        )
        
        print("Link 'Saída/Permanência' encontrado. Clicando no link...")
        saida_permanencia_link.click()

        time.sleep(5)
        print("Página de Saída/Permanência acessada com sucesso!")

        # Mudança de foco para o iframe correto
        print("Tentando mudar o foco para o iframe...")
        WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.NAME, 'f_principal')))
        print("Foco alterado para o iframe.")

        # Clica no botão "PESQUISAR"
        print("Tentando localizar o botão PESQUISAR dentro do iframe...")
        pesquisar_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@name='pesquisar' and @value='PESQUISAR']"))
        )
        
        print("Botão PESQUISAR encontrado!")
        pesquisar_button.click()
        print("Botão PESQUISAR clicado!")

        time.sleep(5)

        # Extração de dados
        nomes = []
        while True:
            # Localiza as linhas da tabela com os dados
            linhas = driver.find_elements(By.XPATH, "//tr[contains(@class, 'linha_selecionavel')]")

            for linha in linhas:
                # Extrai o nome do segundo <td> dentro de cada linha
                nome = linha.find_element(By.XPATH, './td[2]').text
                nomes.append(nome)

            # Tenta localizar o botão "Próxima página"
            try:
                print("Tentando localizar a seta para a próxima página...")
                next_page_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'exibirPagina')]/img[@alt='Proxima']"))
                )
                print("Seta de próxima página encontrada. Clicando na seta...")
                next_page_button.click()
                time.sleep(5)  # Aguarda carregar a próxima página
            except:
                # Se não encontrar o botão "Próxima página", encerra o loop
                print("Não há mais páginas.")
                break

        # Cria um DataFrame com os nomes extraídos
        df = pd.DataFrame(nomes, columns=["Nome"])

        # Salva os dados em uma planilha CSV
        df.to_csv('internados_sisreg.csv', index=False)
        print("Dados salvos em 'internados_sisreg.csv'.")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        driver.quit()

### Definições Interhosp.py
def ler_credenciais_ghosp():
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    usuario_ghosp = config['G-HOSP']['usuario']
    senha_ghosp = config['G-HOSP']['senha']
    
    return usuario_ghosp, senha_ghosp

# Função para encontrar o arquivo mais recente na pasta de Downloads
def encontrar_arquivo_recente(diretorio):
    arquivos = [os.path.join(diretorio, f) for f in os.listdir(diretorio) if os.path.isfile(os.path.join(diretorio, f))]
    arquivos.sort(key=os.path.getmtime, reverse=True)  # Ordena por data de modificação (mais recente primeiro)
    if arquivos:
        return arquivos[0]  # Retorna o arquivo mais recente
    return None

# Função para verificar se a linha contém um nome válido
def linha_valida(linha):
    # Verifica se a primeira coluna tem um número de 6 dígitos
    if re.match(r'^\d{6}$', str(linha[0])):
        # Verifica se a segunda ou a terceira coluna contém um nome válido
        if isinstance(linha[1], str) and linha[1].strip():
            return 'coluna_2'
        elif isinstance(linha[2], str) and linha[2].strip():
            return 'coluna_3'
    return False

# Função principal para extrair os nomes
def extrair_nomes(original_df):
    # Lista para armazenar os nomes extraídos
    nomes_extraidos = []
    
    # Percorre as linhas do DataFrame original
    for i, row in original_df.iterrows():
        coluna = linha_valida(row)
        if coluna == 'coluna_2':
            nome = row[1].strip()  # Extrai da segunda coluna
            nomes_extraidos.append(nome)
        elif coluna == 'coluna_3':
            nome = row[2].strip()  # Extrai da terceira coluna
            nomes_extraidos.append(nome)
        else:
            print(f"Linha ignorada: {row}")
    
    # Converte a lista de nomes extraídos para um DataFrame
    nomes_df = pd.DataFrame(nomes_extraidos, columns=['Nome'])
    
    # Caminho para salvar o novo arquivo sobrescrevendo o anterior na pasta atual
    caminho_novo_arquivo = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'internados_ghosp.csv')
    nomes_df.to_csv(caminho_novo_arquivo, index=False)
    
    print(f"Nomes extraídos e salvos em {caminho_novo_arquivo}.")

def internhosp():
    usuario, senha = ler_credenciais_ghosp()

    # Caminho para o ChromeDriver
    chrome_driver_path = "chromedriver.exe"
    # Obtém o caminho da pasta de Downloads do usuário
    pasta_downloads = str(Path.home() / "Downloads")

    print(f"Pasta de Downloads: {pasta_downloads}")

    # Inicializa o navegador (Chrome neste caso) usando o serviço
    service = Service(executable_path=chrome_driver_path)
    
    # Inicializa o navegador (Chrome neste caso) usando o serviço
    driver = webdriver.Chrome(service=service)

    # Minimizando a janela após iniciar o Chrome
    driver.minimize_window()

    # Acesse a página de login do G-HOSP
    driver.get('http://10.16.9.43:4001/users/sign_in')

    try:
        # Ajustar o zoom para 50% antes do login
        print("Ajustando o zoom para 50%...")
        driver.execute_script("document.body.style.zoom='50%'")
        time.sleep(2)  # Aguarda um pouco após ajustar o zoom

        # Realiza o login
        print("Tentando localizar o campo de e-mail...")
        email_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "email"))
        )
        email_field.send_keys(usuario)

        print("Tentando localizar o campo de senha...")
        senha_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "user_password"))
        )
        senha_field.send_keys(senha)

        print("Tentando localizar o botão de login...")
        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//input[@value='Entrar']"))
        )
        login_button.click()

        time.sleep(5)
        print("Login realizado com sucesso!")

        # Acessar a página de relatórios
        print("Acessando a página de relatórios...")
        driver.get('http://10.16.9.43:4001/relatorios/rc001s')

        # Ajustar o zoom para 60% após acessar a página de relatórios
        print("Ajustando o zoom para 60% na página de relatórios...")
        driver.execute_script("document.body.style.zoom='60%'")
        time.sleep(2)  # Aguarda um pouco após ajustar o zoom

        # Selecionar todas as opções no dropdown "Setor"
        print("Selecionando todos os setores...")
        setor_select = Select(WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "setor_id1"))
        ))
        for option in setor_select.options:
            print(f"Selecionando o setor: {option.text}")  # Para garantir que todos os setores estão sendo selecionados
            setor_select.select_by_value(option.get_attribute('value'))

        print("Todos os setores selecionados!")

        # Maximiza a janela para garantir que todos os elementos estejam visíveis
        driver.maximize_window()
        
        # Selecionar o formato CSV
        print("Rolando até o dropdown de formato CSV...")
        formato_dropdown = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "tipo_arquivo"))
        )
        driver.execute_script("arguments[0].scrollIntoView(true);", formato_dropdown)
        time.sleep(2)

        print("Selecionando o formato CSV...")
        formato_select = Select(formato_dropdown)
        formato_select.select_by_value("csv")

        print("Formato CSV selecionado!")

        # Clicar no botão "Imprimir"
        print("Tentando clicar no botão 'IMPRIMIR'...")
        imprimir_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "enviar_relatorio"))
        )
        imprimir_button.click()

        print("Relatório sendo gerado!")

        driver.minimize_window()

        # Aguardar até que o arquivo CSV seja baixado
        while True:
            arquivo_recente = encontrar_arquivo_recente(pasta_downloads)
            if arquivo_recente and arquivo_recente.endswith('.csv'):
                print(f"Arquivo CSV encontrado: {arquivo_recente}")
                break
            else:
                print("Aguardando o download do arquivo CSV...")
                time.sleep(5)  # Aguarda 5 segundos antes de verificar novamente

        try:
            # Carregar o arquivo CSV recém-baixado, garantindo que todas as colunas sejam lidas como texto
            original_df = pd.read_csv(arquivo_recente, header=None, dtype=str)

            # Chamar a função para extrair os nomes do CSV recém-baixado
            extrair_nomes(original_df)

        except Exception as e:
            print(f"Erro ao processar o arquivo CSV: {e}")

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
    finally:
        driver.quit()

def trazer_terminal():
    # Obtenha o identificador da janela do terminal
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    hwnd = kernel32.GetConsoleWindow()
    
    if hwnd != 0:
        user32.ShowWindow(hwnd, 5)  # 5 = SW_SHOW (Mostra a janela)
        user32.SetForegroundWindow(hwnd)  # Traz a janela para frente

### Funções do motivo_alta.py

def motivo_alta():
    # Função para ler a lista de pacientes de alta do CSV
    def ler_pacientes_de_alta():
        df = pd.read_csv('pacientes_de_alta.csv')
        return df

    # Função para salvar a lista com o motivo de alta
    def salvar_pacientes_com_motivo(df):
        df.to_csv('pacientes_de_alta.csv', index=False)

    # Inicializa o ChromeDriver
    def iniciar_driver():
        chrome_driver_path = "chromedriver.exe"
        service = Service(executable_path=chrome_driver_path)
        driver = webdriver.Chrome(service=service)
        driver.maximize_window()
        return driver

    # Função para realizar login no G-HOSP
    def login_ghosp(driver, usuario, senha):
        driver.get('http://10.16.9.43:4002/users/sign_in')

        # Ajusta o zoom para 50%
        driver.execute_script("document.body.style.zoom='50%'")
        time.sleep(2)
        trazer_terminal()
        
        # Localiza os campos visíveis de login
        email_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "email")))
        email_field.send_keys(usuario)
        
        senha_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "password")))
        senha_field.send_keys(senha)
        
        # Atualiza os campos ocultos com os valores corretos e simula o clique no botão de login
        driver.execute_script("""
            document.getElementById('user_email').value = arguments[0];
            document.getElementById('user_password').value = arguments[1];
            document.getElementById('new_user').submit();
        """, usuario, senha)

    # Função para pesquisar um nome e obter o motivo de alta via HTML
    def obter_motivo_alta(driver, nome):
        driver.get('http://10.16.9.43:4002/prontuarios')

        # Localiza o campo de nome e insere o nome do paciente
        nome_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "nome")))
        nome_field.send_keys(nome)
        
        # Clica no botão de procurar
        procurar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//input[@value='Procurar']")))
        procurar_button.click()

        # Aguarda a página carregar
        time.sleep(10)
        
        try:
            # Localiza o elemento com o rótulo "Motivo da alta"
            motivo_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//small[text()='Motivo da alta: ']"))
            )

            # Agora captura o conteúdo do próximo elemento <div> após o rótulo
            motivo_conteudo_element = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//small[text()='Motivo da alta: ']/following::div[@class='pl5 pb5']"))
            )
            
            motivo_alta = motivo_conteudo_element.text
            print(f"Motivo de alta capturado: {motivo_alta}")
            
        except Exception as e:
            motivo_alta = "Motivo da alta não encontrado"
            print(f"Erro ao capturar motivo da alta para {nome}: {e}")
        
        return motivo_alta

    # Função principal para processar a lista de pacientes e buscar o motivo de alta
    def processar_lista():
        usuario, senha = ler_credenciais_ghosp()
        driver = iniciar_driver()
        
        # Faz login no G-HOSP
        login_ghosp(driver, usuario, senha)
        
        # Lê a lista de pacientes de alta
        df_pacientes = ler_pacientes_de_alta()
        
        # Verifica cada paciente e adiciona o motivo de alta
        for i, row in df_pacientes.iterrows():
            nome = row['Nome']
            print(f"Buscando motivo de alta para: {nome}")
            
            motivo = obter_motivo_alta(driver, nome)
            df_pacientes.at[i, 'Motivo da Alta'] = motivo
            print(f"Motivo de alta para {nome}: {motivo}")
            
            time.sleep(2)  # Tempo de espera entre as requisições

        # Salva o CSV atualizado
        salvar_pacientes_com_motivo(df_pacientes)
        
        driver.quit()

    # Execução do script
    if __name__ == '__main__':
        processar_lista()

# Função principal
def main():
    while True:
        limpar_tela()
        exibir_cabecalho()
        exibir_menu()

        opcao = input("\nEscolha uma opção: ")

        if opcao == '1':
            limpar_tela()
            print(Fore.GREEN + "Executando extração dos internados SISREG...")
            extrator()
        elif opcao == '2':
            limpar_tela()
            print(Fore.GREEN + "Executando extração dos internados G-HOSP...")
            internhosp()
        elif opcao == '3':
            limpar_tela()
            print(Fore.YELLOW + "Comparando e tratando dados...")
            comparar_dados()
        elif opcao == '4':
            limpar_tela()
            print(Fore.GREEN + "Capturando motivos de alta...")
            motivo_alta()
        elif opcao == '5':
            limpar_tela()
            print(Fore.CYAN + "Encerrando o programa. Até mais!")
            break
        else:
            print(Fore.RED + "Opção inválida! Tente novamente.")
            time.sleep(2)

##if __name__ == '__main__':
##    main()

######################
## INÍCIO DA CODIFICAÇÃO DA INTERFAÇE GRAFICA
######################

# Função para redirecionar a saída do terminal para a Text Box
class RedirectOutputToGUI:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)  # Auto scroll para o final da Text Box

    def flush(self):
        pass

# Função para normalizar o nome (remover acentos, transformar em minúsculas)
def normalizar_nome(nome):
    nfkd = unicodedata.normalize('NFKD', nome)
    return "".join([c for c in nfkd if not unicodedata.combining(c)]).lower()

# Função para comparar os arquivos CSV e salvar os pacientes a dar alta
def comparar_dados():
    print("Comparando dados...")
    arquivo_sisreg = 'internados_sisreg.csv'
    arquivo_ghosp = 'internados_ghosp.csv'

    if not os.path.exists(arquivo_sisreg) or not os.path.exists(arquivo_ghosp):
        print("Os arquivos internados_sisreg.csv ou internados_ghosp.csv não foram encontrados!")
        return

    with open(arquivo_sisreg, 'r', encoding='utf-8') as sisreg_file:
        sisreg_nomes_lista = [normalizar_nome(linha[0].strip()) for linha in csv.reader(sisreg_file) if linha]

    sisreg_nomes = set(sisreg_nomes_lista[1:])

    with open(arquivo_ghosp, 'r', encoding='utf-8') as ghosp_file:
        ghosp_nomes = {normalizar_nome(linha[0].strip()) for linha in csv.reader(ghosp_file) if linha}

    pacientes_a_dar_alta = sisreg_nomes - ghosp_nomes

    if pacientes_a_dar_alta:
        print("\n---===> PACIENTES A DAR ALTA <===---")
        for nome in sorted(pacientes_a_dar_alta):
            print(nome)

        with open('pacientes_de_alta.csv', 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Nome'])
            for nome in sorted(pacientes_a_dar_alta):
                writer.writerow([nome])

        print("\nA lista de pacientes a dar alta foi salva em 'pacientes_de_alta.csv'.")
    else:
        print("\nNenhum paciente a dar alta encontrado!")

# Função para executar o extrator e atualizar o status na interface
def executar_sisreg():
    def run_task():
        try:
            extrator()
            messagebox.showinfo("Sucesso", "Extração dos internados SISREG realizada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    threading.Thread(target=run_task).start()  # Executar a função em um thread separado

# Função para executar a extração do G-HOSP
def executar_ghosp():
    def run_task():
        try:
            internhosp()
            messagebox.showinfo("Sucesso", "Extração dos internados G-HOSP realizada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    threading.Thread(target=run_task).start()

# Função para comparar os dados
def comparar():
    def run_task():
        try:
            comparar_dados()
            messagebox.showinfo("Sucesso", "Comparação de dados realizada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    threading.Thread(target=run_task).start()

# Função para capturar o motivo de alta
def capturar_motivo_alta():
    def run_task():
        try:
            motivo_alta()
            messagebox.showinfo("Sucesso", "Motivos de alta capturados com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
    threading.Thread(target=run_task).start()

# Função para abrir e editar o arquivo config.ini
def abrir_configuracoes():
    def salvar_configuracoes():
        try:
            with open('config.ini', 'w') as configfile:
                configfile.write(text_area.get("1.0", tk.END))
            messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {e}")

    # Cria uma nova janela para editar o arquivo config.ini
    janela_config = tk.Toplevel()
    janela_config.title("Editar Configurações - config.ini")
    janela_config.geometry("500x400")

    # Área de texto para exibir e editar o conteúdo do config.ini
    text_area = scrolledtext.ScrolledText(janela_config, wrap=tk.WORD, width=60, height=20)
    text_area.pack(pady=10, padx=10)

    try:
        with open('config.ini', 'r') as configfile:
            text_area.insert(tk.END, configfile.read())
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo config.ini não encontrado!")

    # Botão para salvar as alterações
    btn_salvar = tk.Button(janela_config, text="Salvar", command=salvar_configuracoes)
    btn_salvar.pack(pady=10)

# URL do JSON com as versões e links de download do ChromeDriver
CHROMEDRIVER_VERSIONS_URL = "https://googlechromelabs.github.io/chrome-for-testing/known-good-versions-with-downloads.json"

# Função para obter a versão do Google Chrome
def obter_versao_chrome():
    try:
        print("Verificando a versão do Google Chrome instalada...")
        process = subprocess.run(
            ['reg', 'query', 'HKEY_CURRENT_USER\\Software\\Google\\Chrome\\BLBeacon', '/v', 'version'],
            capture_output=True,
            text=True
        )
        version_line = process.stdout.strip().split()[-1]
        print(f"Versão do Google Chrome encontrada: {version_line}")
        return version_line
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao obter a versão do Google Chrome: {e}")
        print(f"Erro ao obter a versão do Google Chrome: {e}")
        return None

# Função para obter a versão do ChromeDriver
def obter_versao_chromedriver():
    try:
        print("Verificando a versão atual do ChromeDriver...")
        process = subprocess.run(
            ['chromedriver', '--version'],
            capture_output=True,
            text=True
        )
        version_line = process.stdout.strip().split()[1]
        print(f"Versão do ChromeDriver encontrada: {version_line}")
        return version_line
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao obter a versão do ChromeDriver: {e}")
        print(f"Erro ao obter a versão do ChromeDriver: {e}")
        return None

# Função para consultar o JSON e obter o link de download da versão correta do ChromeDriver
def buscar_versao_chromedriver(versao_chrome):
    try:
        print(f"Buscando a versão compatível do ChromeDriver para o Google Chrome {versao_chrome}...")
        response = requests.get(CHROMEDRIVER_VERSIONS_URL)
        if response.status_code != 200:
            messagebox.showerror("Erro", f"Erro ao acessar o JSON de versões: Status {response.status_code}")
            print(f"Erro ao acessar o JSON de versões: Status {response.status_code}")
            return None
        
        major_version = versao_chrome.split('.')[0]
        json_data = response.json()
        for version_data in json_data["versions"]:
            if version_data["version"].startswith(major_version):
                for download in version_data["downloads"]["chromedriver"]:
                    if "win32" in download["url"]:
                        print(f"Versão do ChromeDriver encontrada: {version_data['version']}")
                        return download["url"]
        
        messagebox.showerror("Erro", f"Não foi encontrada uma versão correspondente do ChromeDriver para a versão {versao_chrome}")
        print(f"Não foi encontrada uma versão correspondente do ChromeDriver para o Google Chrome {versao_chrome}")
        return None
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o JSON do ChromeDriver: {e}")
        print(f"Erro ao processar o JSON do ChromeDriver: {e}")
        return None

# Função para baixar o ChromeDriver
def baixar_chromedriver(chromedriver_url):
    try:
        print(f"Baixando o ChromeDriver de {chromedriver_url}...")
        response = requests.get(chromedriver_url, stream=True)
        
        if response.status_code != 200:
            messagebox.showerror("Erro", f"Não foi possível baixar o ChromeDriver: Status {response.status_code}")
            print(f"Não foi possível baixar o ChromeDriver: Status {response.status_code}")
            return
        
        # Salva o arquivo ZIP do ChromeDriver
        with open("chromedriver_win32.zip", "wb") as file:
            for chunk in response.iter_content(chunk_size=8192):
                file.write(chunk)
                print(".", end="", flush=True)  # Imprime pontos para acompanhar o progresso
        print("\nDownload concluído. Extraindo o arquivo ZIP...")
        
        # Extrai o conteúdo do ZIP
        with zipfile.ZipFile("chromedriver_win32.zip", 'r') as zip_ref:
            zip_ref.extractall(".")  # Extrai para a pasta atual

        # Descobre o diretório onde o script está rodando
        pasta_atual = os.path.dirname(os.path.abspath(__file__))
        
        # Caminho para o ChromeDriver extraído
        chromedriver_extraido = os.path.join(pasta_atual, "chromedriver-win32", "chromedriver.exe")
        destino_chromedriver = os.path.join(pasta_atual, "chromedriver.exe")

        if os.path.exists(chromedriver_extraido):
            print(f"Movendo o ChromeDriver para {destino_chromedriver}...")
            shutil.move(chromedriver_extraido, destino_chromedriver)
            print("Atualização do ChromeDriver concluída!")
        else:
            print(f"Erro: o arquivo {chromedriver_extraido} não foi encontrado.")

        # Remove o arquivo ZIP após a extração
        os.remove("chromedriver_win32.zip")
        
        messagebox.showinfo("Sucesso", "ChromeDriver atualizado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao atualizar o ChromeDriver: {e}")
        print(f"Erro ao atualizar o ChromeDriver: {e}")

# Função para verificar a versão do Chrome e ChromeDriver e atualizar, se necessário
def verificar_atualizar_chromedriver():
    versao_chrome = obter_versao_chrome()
    versao_chromedriver = obter_versao_chromedriver()
    
    if versao_chrome and versao_chromedriver:
        if versao_chrome.split('.')[0] == versao_chromedriver.split('.')[0]:
            print("Versão do ChromeDriver e Google Chrome são compatíveis.")
            messagebox.showinfo("Versões Compatíveis", f"Versão do Chrome ({versao_chrome}) e ChromeDriver ({versao_chromedriver}) são compatíveis.")
        else:
            resposta = messagebox.askyesno("Atualização Necessária", f"A versão do ChromeDriver ({versao_chromedriver}) não é compatível com o Chrome ({versao_chrome}). Deseja atualizar?")
            if resposta:
                chromedriver_url = buscar_versao_chromedriver(versao_chrome)
                if chromedriver_url:
                    baixar_chromedriver(chromedriver_url)

def mostrar_versao():
    versao = "AUTOMATOR - AUTOREG\nAutomação de Pacientes a Dar Alta - SISREG & G-HOSP\nVersão 2.5 - Outubro de 2024\nAutor: Michel R. Paes\nGithub: MrPaC6689\nDesenvolvido com o apoio do ChatGPT\nContato: michelrpaes@gmail.com"
    messagebox.showinfo("Automator 2.5", versao)

# Função para exibir o conteúdo do arquivo README.md
def exibir_leia_me():
    try:
        # Verifica se o arquivo README.md existe
        if not os.path.exists('README.md'):
            messagebox.showerror("Erro", "O arquivo README.md não foi encontrado.")
            return
        
        # Lê o conteúdo do arquivo README.md
        with open('README.md', 'r', encoding='utf-8') as file:
            conteudo = file.read()
        
        # Cria uma nova janela para exibir o conteúdo
        janela_leia_me = tk.Toplevel()
        janela_leia_me.title("Leia-me")
        janela_leia_me.geometry("1000x800")
        
        # Cria uma área de texto com scroll para exibir o conteúdo
        text_area = scrolledtext.ScrolledText(janela_leia_me, wrap=tk.WORD, width=120, height=45)
        text_area.pack(pady=10, padx=10)
        text_area.insert(tk.END, conteudo)
        text_area.config(state=tk.DISABLED)  # Desabilita a edição do texto

         # Adiciona um botão "Fechar" para fechar a janela do Leia-me
        btn_fechar = tk.Button(janela_leia_me, text="Fechar", command=janela_leia_me.destroy)
        btn_fechar.pack(pady=10)
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao tentar abrir o arquivo README.md: {e}")

# Função para abrir o arquivo CSV com o programa de planilhas padrão
def abrir_csv(caminho_arquivo):
    try:
        if os.path.exists(caminho_arquivo):
            if os.name == 'nt':  # Windows
                print("Abrindo o arquivo CSV como planilha, aguarde...")
                os.startfile(caminho_arquivo)              
            elif os.name == 'posix':  # macOS ou Linux
                print("Abrindo o arquivo CSV como planilha, aguarde...")
                subprocess.call(('xdg-open' if 'linux' in os.sys.platform else 'open', caminho_arquivo))
        else:
            print("O arquivo {caminho_arquivo} não foi encontrado.")
            messagebox.showerror("Erro", f"O arquivo {caminho_arquivo} não foi encontrado.")            
    except Exception as e:
        print("Não foi possível abrir o arquivo: {e}")
        messagebox.showerror("Erro", f"Não foi possível abrir o arquivo: {e}")

# Função para sair do programa
def sair_programa():
    janela.destroy()
    
# Função principal da interface gráfica
def criar_interface():
    # Cria a janela principal
    global janela  # Declara a variável 'janela' como global para ser acessada em outras funções
    janela = tk.Tk()
    janela.title("Automator - AutoReg - v.2.5 ")
    # Adiciona texto explicativo ou outro conteúdo abaixo do título principal
    tk.Label(janela, text="Automator - AutoReg", font=("Helvetica", 16)).pack(pady=10)
    tk.Label(janela, text="Sistema automatizado para captura de pacientes a dar alta - SISREG G-HOSP.\n Por Michel R. Paes - Outubro 2024 \n Escolha uma opção abaixo:", 
             font=("Helvetica", 12)).pack(pady=10)
    janela.geometry("950x800")

     # Criação do menu
    menubar = tk.Menu(janela)
    janela.config(menu=menubar)

    # Adiciona um submenu "Configurações"
    config_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Configurações", menu=config_menu)
    config_menu.add_command(label="Editar config.ini", command=abrir_configuracoes)
    config_menu.add_command(label="Verificar e Atualizar ChromeDriver", command=verificar_atualizar_chromedriver)

    # Adiciona um submenu "Informações" com "Versão" e "Leia-me"
    info_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Informações", menu=info_menu)
    info_menu.add_command(label="Versão", command=mostrar_versao)
    info_menu.add_command(label="Leia-me", command=exibir_leia_me)

     # Frame para manter os botões lado a lado e padronizar tamanho
    button_width = 25  # Define uma largura fixa para todos os botões

    # Frame para SISREG
    frame_sisreg = tk.Frame(janela)
    frame_sisreg.pack(pady=5)
    
    btn_sisreg = tk.Button(frame_sisreg, text="Extrair internados SISREG", command=executar_sisreg, width=button_width)
    btn_sisreg.pack(side=tk.LEFT, padx=5)
    
    btn_exibir_sisreg = tk.Button(frame_sisreg, text="Exibir Resultado SISREG", command=lambda: abrir_csv('internados_sisreg.csv'), width=button_width)
    btn_exibir_sisreg.pack(side=tk.LEFT, padx=5)

    # Frame para G-HOSP
    frame_ghosp = tk.Frame(janela)
    frame_ghosp.pack(pady=5)
    
    btn_ghosp = tk.Button(frame_ghosp, text="Extrair internados G-HOSP", command=executar_ghosp, width=button_width)
    btn_ghosp.pack(side=tk.LEFT, padx=5)

    btn_exibir_ghosp = tk.Button(frame_ghosp, text="Exibir Resultado G-HOSP", command=lambda: abrir_csv('internados_ghosp.csv'), width=button_width)
    btn_exibir_ghosp.pack(side=tk.LEFT, padx=5)
    
    # Frame para Comparação
    frame_comparar = tk.Frame(janela)
    frame_comparar.pack(pady=5)
    
    btn_comparar = tk.Button(frame_comparar, text="Comparar e tratar dados", command=comparar, width=button_width)
    btn_comparar.pack(side=tk.LEFT, padx=5)
    
    btn_exibir_comparar = tk.Button(frame_comparar, text="Exibir Resultado Comparação", command=lambda: abrir_csv('pacientes_de_alta.csv'), width=button_width)
    btn_exibir_comparar.pack(side=tk.LEFT, padx=5)

    # Botão de Sair
    btn_sair = tk.Button(janela, text="Sair", command=sair_programa, width=2*button_width + 10)  # Largura ajustada para ficar mais largo
    btn_sair.pack(pady=20)

    # Widget de texto com scroll para mostrar o status
    text_area = ScrolledText(janela, wrap=tk.WORD, height=30, width=110)
    text_area.pack(pady=10)

    # Redireciona a saída do terminal para a Text Box
    sys.stdout = RedirectOutputToGUI(text_area)
   
    # Inicia o loop da interface gráfica
    janela.mainloop()
    
    # Executa a interface gráfica
if __name__ == '__main__':
    criar_interface()