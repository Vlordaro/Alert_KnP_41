import os
import datetime
import time
import asyncio
from telegram import Bot
import win32evtlog
import socket

# Configurações iniciais
chat_id = "ID do seu Chat do telegram"
token = "Token ID do seu Bot telegram"
bot = Bot(token=token)
log_folder_path = r"caminho do arquivo de Logs"
indice_arquivo = os.path.join(log_folder_path, "ultimo_indice.txt")

# Obter o nome do computador
computer_name = socket.gethostname()

# Nome do arquivo de log com o nome do computador
log_filename = f"Kernel_PnP_41_log_{computer_name}.txt"

# Criação da pasta de logs, se não existir
os.makedirs(log_folder_path, exist_ok=True)

async def enviar_mensagem_telegram(mensagem):
    tentativas = 0
    max_tentativas = 3
    enviado = False
    while tentativas < max_tentativas and not enviado:
        try:
            await bot.send_message(chat_id=chat_id, text=mensagem)
            enviado = True
        except Exception as e:
            error_message = str(e)
            print(f"Erro ao enviar mensagem: {error_message}")
            tentativas += 1
            time.sleep(10)

def registrar_log(timestamp, evento, maquina):
    log_file_path = os.path.join(log_folder_path, log_filename)
    try:
        with open(log_file_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"{timestamp} - {evento} - {maquina}\n")
        print(f"Log registrado em: {log_file_path}")
    except Exception as e:
        print(f"Erro ao registrar log: {str(e)}")

def obter_ultimo_indice():
    try:
        with open(indice_arquivo, "r") as f:
            return int(f.read().strip())
    except FileNotFoundError:
        print(f"Arquivo de índice não encontrado. Criando novo arquivo em: {indice_arquivo}")
        salvar_ultimo_indice(0)
        return 0
    except ValueError:
        print("Erro ao ler o último índice. Iniciando do zero.")
        salvar_ultimo_indice(0)
        return 0

def salvar_ultimo_indice(indice):
    try:
        with open(indice_arquivo, "w") as f:
            f.write(str(indice))
        print(f"Índice {indice} salvo em: {indice_arquivo}")
    except Exception as e:
        print(f"Erro ao salvar índice: {str(e)}")

def monitorar_eventos():
    server = 'localhost'
    logtype = 'System'
    hand = win32evtlog.OpenEventLog(server, logtype)
    ultimo_indice = obter_ultimo_indice()
    
    print(f"Iniciando monitoramento. Último índice: {ultimo_indice}")
    
    while True:
        flags = win32evtlog.EVENTLOG_BACKWARDS_READ | win32evtlog.EVENTLOG_SEQUENTIAL_READ
        events = win32evtlog.ReadEventLog(hand, flags, 0)
        
        for event in events:
            if event.RecordNumber > ultimo_indice and event.EventID == 41 and "Microsoft-Windows-Kernel-Power" in event.SourceName:
                current_time = datetime.datetime.now().strftime("%d-%m-%Y %p %H:%M:%S")
                
                # Mensagem para o Telegram
                mensagem_telegram = (f"⚠️ Alerta de Kernel PnP 41\n\n"
                                     f"Um evento do Kernel-Power (ID 41) foi registrado na loja {computer_name} "
                                     f"em {current_time}.")
                
                # Registro simplificado para o arquivo de log
                registrar_log(current_time, "Kernel Power 41", computer_name)
                
                asyncio.run(enviar_mensagem_telegram(mensagem_telegram))
                
                ultimo_indice = event.RecordNumber
                salvar_ultimo_indice(ultimo_indice)
        
        time.sleep(30)

if __name__ == "__main__":
    monitorar_eventos()
