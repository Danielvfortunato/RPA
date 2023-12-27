import subprocess
import time
import pyautogui
import requests
import io
import database
import os
from datetime import datetime
import traceback

def capture_screen():
    screenshot = pyautogui.screenshot()
    screenshot_bytes = io.BytesIO()
    screenshot.save(screenshot_bytes, format='PNG')
    screenshot_bytes.seek(0)
    return screenshot_bytes

# def send_message_with_screenshot(error_message, screenshot_bytes):
#     chat_ids_results = database.consultar_chat_id_pamela()
#     token_result = database.consultar_token_bot()
#     chat_ids = [result[0] for result in chat_ids_results]
#     token = token_result[0][0]
#     base_url = f"https://api.telegram.org/bot{token}/sendPhoto"
#     for chat_id in chat_ids:
#         files = {'photo': screenshot_bytes}
#         data = {'chat_id': chat_id, 'caption': error_message}
#         response = requests.post(base_url, files=files, data=data)
#         if response.status_code != 200:
#             print(f"Failed to send message to chat_id {chat_id}. Response: {response.content}")
#         else:
#             print(f"Message sent successfully to chat_id {chat_id} on Telegram!")

def send_message_with_screenshot(error_message, screenshot_bytes):
    chat_ids_results = database.consultar_chat_id_pamela()
    token_result = database.consultar_token_bot()
    chat_ids = [result[0] for result in chat_ids_results]
    token = token_result[0][0]
    base_url = f"https://api.telegram.org/bot{token}/sendPhoto"
    for chat_id in chat_ids:
        # Resetar o cursor para o início do stream antes de cada envio
        screenshot_bytes.seek(0)
        files = {'photo': screenshot_bytes}
        data = {'chat_id': chat_id, 'caption': error_message}
        response = requests.post(base_url, files=files, data=data)
        if response.status_code != 200:
            print(f"Failed to send message to chat_id {chat_id}. Response: {response.content}")
        else:
            print(f"Message sent successfully to chat_id {chat_id} on Telegram!")

def check_if_process_running(process_command):
    for line in os.popen('wmic process get description, commandline'):
        if process_command in line:
            return True
    return False

def is_within_blocked_time():
    current_time = datetime.now().time()
    blocked_times = [("10:20", "10:35"), ("12:20", "12:35"), ("17:20", "17:35"), ("22:20", "22:35")]

    for start, end in blocked_times:
        start_time = datetime.strptime(start, "%H:%M").time()
        end_time = datetime.strptime(end, "%H:%M").time()

        if start_time <= current_time <= end_time:
            return True
    return False


def run_script():
    if not check_if_process_running('main.py'):
        print('main.py is not running. Attempting to start it.')

        if check_if_process_running('SisFin.exe'):
            subprocess.run(['taskkill', '/F', '/IM', 'SisFin.exe'])
            time.sleep(3)
        if check_if_process_running('chrome.exe'):
            subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'])
            time.sleep(3)

        process = subprocess.Popen(['python', r'C:\Users\user\Documents\rpa_Project\main.py'], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        stdout, stderr = process.communicate()

        if process.returncode != 0:
            try:
                # Open the file with 'replace' error handler to manage encoding errors
                with open('info.txt', 'r', encoding='utf-8', errors='replace') as file:
                    line = file.readline()
                    id_solicitacao, numerodocto = line.strip().split(',')

                # Decode the stderr and extract the text after "traceback"
                stderr_decoded = stderr.decode('utf-8', errors='replace')
                traceback_index = stderr_decoded.find('Traceback (most recent call last):')
                if traceback_index != -1:
                    error_details = stderr_decoded[traceback_index:traceback_index + 500]
                else:
                    error_details = stderr_decoded[:500]

                error_message = (f"Erro detectado em main.py:\nID Solicitacao: {id_solicitacao}\nNumero Docto: {numerodocto}\n" +
                                 f"{error_details}")
                screenshot = capture_screen()
                send_message_with_screenshot(error_message, screenshot)
            except Exception as e:
                print(f"Erro ao ler o arquivo de informações: {e}")

    else:
        print('main.py is already running. Doing nothing.')


if __name__ == "__main__":
    while True:
        run_script()
        time.sleep(10)
