import pandas as pd
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import os

# Redirecionar a saída de erro para o dispositivo que descarta tudo
import sys
sys.stderr = open(os.devnull, "w")

# Função para agendar reunião
def agendar_reuniao(destinatario, assunto, corpo, data, hora, duracao, lembrete, local):
    try:
        outlook = win32.Dispatch('Outlook.Application')
        reuniao = outlook.CreateItem(1)
        reuniao.MeetingStatus = 1
        reuniao.Subject = assunto
        reuniao.Body = corpo
        
        # Convertendo a data e hora para o fuso horário de São Paulo
        data_hora_str = f"{data} {hora}"
        data_hora_utc = pd.to_datetime(data_hora_str, format='%d/%m/%Y %H:%M:%S').tz_localize('UTC')
        data_hora_sp = data_hora_utc.tz_convert('America/Sao_Paulo')
        
        reuniao.Start = data_hora_sp
        reuniao.Duration = duracao
        reuniao.ReminderSet = True
        reuniao.ReminderMinutesBeforeStart = lembrete
        reuniao.Location = local
        reuniao.Recipients.Add(destinatario)
        reuniao.Send()
        print(f"Reunião marcada com sucesso para {destinatario}.")
    except Exception as e:
        print(f"Erro ao marcar reunião para {destinatario}: {e}")

# Função para fazer upload do arquivo Excel
def upload_excel():
    root = tk.Tk()
    root.withdraw()  # Esconder a janela principal

    input("***NOTA: ANTES DE EXECUTAR O ARQUIVO, LEIA A DOCUMENTAÇÃO *** \nPressione enter e aguarde iniciar...")

    # Abrir o dialogo para seleção do arquivo
    arquivo = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

    if arquivo:
        # Carregar o arquivo Excel
        return pd.read_excel(arquivo)
    else:
        return None

# Função para solicitar informações adicionais do usuário
def solicitar_informacoes():
    while True:
        print('='*60)
        assunto = input("ASSUNTO: ")
        print('='*60)
        corpo_email = input("CORPO E-MAIL: ")
        print('='*60)
        data = validar_formato_data("DATA DA REUNIÃO (dd/mm/aaaa): ")
        print('='*60)
        duracao = validar_formato_numero("DURAÇÃO DA REUINÃO (em minutos): ")
        print('='*60)
        lembrete = validar_formato_numero("LEMBRETE (em minutos): ")
        print('='*60)
        local = input("LOCAL: ")
        print('='*60)

        if data and duracao and lembrete:
            return assunto, corpo_email, data, duracao, lembrete, local

# Função para validar o formato da data
def validar_formato_data(mensagem):
    while True:
        data = input(mensagem)
        try:
            pd.to_datetime(data, format='%d/%m/%Y')
            return data
        except ValueError:
            print("Formato de data inválido. Utilize o formato 'dd/mm/aaaa'.")

# Função para validar o formato do número
def validar_formato_numero(mensagem):
    while True:
        numero = input(mensagem)
        if numero.isdigit():
            return int(numero)
        else:
            print("Por favor, insira um número inteiro válido.")

# Solicitar upload do arquivo Excel
arquivo = upload_excel()

if arquivo is not None:
    # Solicitar informações adicionais do usuário
    assunto, corpo_email, data, duracao, lembrete, local = solicitar_informacoes()

    # Loop sobre as linhas do arquivo Excel
    for index, linha in arquivo.iterrows():
        hora = linha['Hora']
        email = linha['Email']

        # Verificar se a célula de hora está preenchida e se o email é uma string válida
        if pd.notna(hora) and isinstance(email, str):
            agendar_reuniao(email, assunto, corpo_email, data, hora, duracao, lembrete, local)
else:
    print("Nenhum arquivo selecionado. O programa será encerrado.")

input("Pressione Enter para sair...")
