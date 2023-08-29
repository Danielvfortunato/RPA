import subprocess
import time
import os

def check_if_process_running(process_command):
    for line in os.popen('wmic process get description, commandline'):
        if process_command in line:
            return True
    return False

def run_script():
    if not check_if_process_running('main.py'):
        print('main.py is not running. Attempting to stop SisFin.exe (if exists) and starting main.py')
            
        if check_if_process_running('SisFin.exe'):
            subprocess.run(['taskkill', '/F', '/IM', 'SisFin.exe'])
            time.sleep(5)
        
        if check_if_process_running('chrome.exe'):
            subprocess.run(['taskkill', '/F', '/IM', 'chrome.exe'])
            time.sleep(5)
            
        subprocess.run(['python', r'C:\Users\user\Documents\rpa_Project\main.py'])
    else:
        print('main.py is already running. Doing nothing.')

if __name__ == "__main__":
    run_script()
