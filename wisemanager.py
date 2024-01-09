from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import psutil
import subprocess
import time
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
import pyautogui
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.select import Select
import pygetwindow as gw
from selenium.webdriver.support import expected_conditions as EC
import os
from pathlib import Path
import xml.etree.ElementTree as ET
import requests
import database

class Wise():
    def __init__(self):
        self.driver = None

    @staticmethod
    def janela_esta_visivel(title):
        windows = gw.getWindowsWithTitle(title)
        if windows and windows[0].visible:
            return True
        return False
  
    def esperar_janela_visivel(self, title, timeout=60):
        start_time = time.time()
        while not self.janela_esta_visivel(title):
            if time.time() - start_time > timeout:
                print(f"Falha: Janela {title} não ficou visível após {timeout} segundos.")
                return False
            time.sleep(1)
        return True
    
    def wait_until_interactive(self, ctrl, timeout=60):
        end_time = time.time() + timeout
        while time.time() < end_time:
            if ctrl.is_visible() and ctrl.is_enabled():
                return True
            time.sleep(0.5)
        raise TimeoutError(f"Elemento não ficou interativo após {timeout} segundos")

    def start_chrome_debugger(self):
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            if proc.info['name'] == 'chrome.exe' and '--remote-debugging-port=9222' in proc.info['cmdline']:
                time.sleep(1)
                self.open_google()
                print("Chrome com depurador remoto já está em execução.")
                return
        try:
            subprocess.Popen([
                "C:\Program Files\Google\Chrome\Application\chrome.exe",
                "--remote-debugging-port=9222",
                "--user-data-dir=C:/Selenium/AutomationProfile"
            ])
            time.sleep(5) 
            print("Chrome com depurador remoto iniciado.")
        except FileNotFoundError:
            print("Caminho do Chrome incorreto ou o Chrome não está instalado.")
        except Exception as e:
            print(f"Erro ao iniciar o Chrome: {e}")

    def open_google(self):
        try:
            app = Application(backend='uia').connect(title="Barra de Tarefas")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Barra de Tarefa")
            return 
        janela = app['Barra de Tarefas']
        abrir_google = janela.child_window(title="Google Chrome - 1 executando o windows")
        try:
            self.wait_until_interactive(abrir_google)
        except TimeoutError as e:
            print(str(e))
            return
        abrir_google.click_input()

    def init_instance_chrome(self):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        if not self.driver:
            # service = Service(executable_path='./chromedriver.exe')
            self.driver = webdriver.Chrome(options=chrome_options)
            self.driver.maximize_window()

        return self.driver

    def login(self):
        username_field = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.NAME, "j_username")))
        password_field = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.NAME, "j_password")))
        username_field.send_keys('rpa')
        password_field.send_keys('Da131800!')
        password_field.submit()

    def Anexar_AP(self, id_solicitacao, num_controle, num_docto, serie_value):
        self.start_chrome_debugger()
        self.init_instance_chrome()
        self.driver.get("https://gera.wisemanager.com.br/WiseManagerBI/#/financeiro/efetivacaoSolicitacaoGastoNew")
        self.reload_page()
        # time.sleep(3)
        if self.driver.find_elements(By.NAME, "j_username") and self.driver.find_elements(By.NAME, "j_password"):
            self.login()
        time.sleep(2)
        select_element = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'select.form-control.input-sm[ng-model="filtro.campo"]')))
        select = Select(select_element)
        select.select_by_value('id')
        time.sleep(2)
        input_element = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input.form-control.input-sm.filter-picker-round[ng-model="filtro.texto"]')))
        input_element.send_keys(id_solicitacao)
        time.sleep(2)
        search = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'button.btn.btn-circle.btn-bordered.thick.btn-primary.imt-5')))
        search.click()
        time.sleep(2)
        open_detail = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click, 'abreDetalheSolicitacao')]")))
        open_detail.click()
        time.sleep(2)
        numeros_docto = WebDriverWait(self.driver, 60).until(EC.presence_of_all_elements_located((By.NAME, "numeroDocto")))
        controles = WebDriverWait(self.driver, 60).until(EC.presence_of_all_elements_located((By.NAME, "numeroControle")))
        time.sleep(2)
        serie_elements = WebDriverWait(self.driver, 60).until(EC.presence_of_all_elements_located((By.NAME, "serie")))
        for numero_docto_el, controle_el, serie_el in zip(numeros_docto, controles, serie_elements):
            valor_docto = numero_docto_el.get_attribute('value')
            valor_serie = serie_el.get_attribute('value')
            
            if valor_docto == num_docto and valor_serie == serie_value:
                controle_el.send_keys(num_controle)
        
        time.sleep(2)
        file_input = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.XPATH, "//input[@class='file-input']")))
        self.driver.execute_script("arguments[0].click();", file_input)

    def get_pdf_file(self, numero_docto, id_solicitacao):
        title = "Abrir"
        time.sleep(3)
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de página pegar pdf não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend="win32").connect(title=title)
        except ElementNotFoundError:
            print("Falha ao conectar com página pegar pdf")
            return 
        janela = app[title]
        try:
            get_document = janela.child_window(title="APs (fixo)", found_index=0)
            get_document.click_input()
        except:
            pass
        
        file_name = janela.child_window(class_name="Edit")
        time.sleep(2)
        self.type_slowly(file_name, f"AP_{numero_docto}{id_solicitacao}.PDF")
        # file_name.type_keys(f"AP_{numero_docto}{id_solicitacao}.PDF")
        time.sleep(1)
        file_name.type_keys("{ENTER}")

    def confirm(self):
        url_esperada = "https://gera.wisemanager.com.br/WiseManagerBI/#/financeiro/efetivacaoSolicitacaoGastoNew"
        if self.driver.current_url == url_esperada:
            # self.reload_page()
            controles = WebDriverWait(self.driver, 60).until(EC.presence_of_all_elements_located((By.NAME, "numeroControle")))
            todos_preenchidos = all(controle.get_attribute('value') != "" for controle in controles if controle.is_displayed())
            time.sleep(1)
            pyautogui.scroll(1000)
            time.sleep(2)
            if todos_preenchidos:
                botao_confirmar = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@title, 'Confirma Lançamento no ERP')]")))
                time.sleep(2)
                # self.driver.execute_script("arguments[0].scrollIntoView();", botao_confirmar)
                botao_confirmar.click()
                time.sleep(3)
                confirmar = r"C:\Users\user\Documents\RPA_Project\imagens\Confirmar.PNG"
                time.sleep(3)
                self.click_specific_button_wise(confirmar)
            else:
                print("Nem todos os controles estão preenchidos. Não foi possível confirmar.")
                
    def click_specific_button_wise(self, button_image_path, confidence_level=0.8):
        button_location = pyautogui.locateOnScreen(button_image_path, confidence=confidence_level)
        if button_location:
            button_x, button_y, button_width, button_height = button_location
            button_center_x = button_x + button_width // 2
            button_center_y = button_y + button_height // 2
            pyautogui.click(button_center_x, button_center_y)
            print("Botão clicado!")
        else:
            print("Botão não encontrado.")

    def click_specific_button_wise_1(self, button_image_path, confidence_level=0.9):
        button_location = pyautogui.locateOnScreen(button_image_path, confidence=confidence_level)
        if button_location:
            button_x, button_y, button_width, button_height = button_location
            button_center_x = button_x + button_width // 2
            button_center_y = button_y + button_height // 2
            pyautogui.click(button_center_x, button_center_y)
            print("Botão clicado!")
        else:
            print("Botão não encontrado.")
    
    def get_nf_values(self, num_docto, id_solicitacao):
        self.start_chrome_debugger()
        self.init_instance_chrome()
        time.sleep(2)
        self.driver.get("https://gera.wisemanager.com.br/WiseManagerBI/#/financeiro/efetivacaoSolicitacaoGastoNew")
        self.reload_page()
        if self.driver.find_elements(By.NAME, "j_username") and self.driver.find_elements(By.NAME, "j_password"):
            self.login()
            time.sleep(2)
        select_element = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'select.form-control.input-sm[ng-model="filtro.campo"]')))
        select = Select(select_element)
        select.select_by_value('id')
        time.sleep(2)
        input_element = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'input.form-control.input-sm.filter-picker-round[ng-model="filtro.texto"]')))
        input_element.send_keys(id_solicitacao)
        time.sleep(2)
        search = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn.btn-circle.btn-bordered.thick.btn-primary.imt-5')))
        search.click()
        time.sleep(2)
        open_detail = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.XPATH, "//button[contains(@ng-click, 'abreDetalheSolicitacao')]")))
        open_detail.click()
        time.sleep(2)
        link_elements = WebDriverWait(self.driver, 60).until(EC.presence_of_all_elements_located((By.XPATH, "//a[contains(@ng-show, '.pdf') and contains(translate(@href, 'PDF', 'pdf'), '.pdf') and contains(@href, 'anexo/solicitacaoGasto')]")))
        numeros_docto = WebDriverWait(self.driver, 60).until(
            EC.presence_of_all_elements_located((By.NAME, "numeroDocto"))
        )
        # pyautogui.scroll(-1000) 
        download_buttons = WebDriverWait(self.driver, 60).until(
            EC.presence_of_all_elements_located((
                By.XPATH, 
                "//a[contains(@href, 'anexo/solicitacaoGasto') and @ng-show=\"anexoSolicitacaoGasto.nomeArquivo.endsWith('.xml')\"]"
            ))
        )
        download_buttons_visible = any(button.is_displayed() for button in download_buttons)
        # self.driver.execute_script("arguments[0].scrollIntoView();", download_buttons_visible)
        for numero_docto, link_element in zip(numeros_docto, link_elements):
            valor_docto = numero_docto.get_attribute('value')
            if str(valor_docto) == str(num_docto) and link_element.is_displayed():
                link_element.click()
                break
        time.sleep(4)
        fechar_popup = r"C:\Users\user\Documents\RPA_Project\imagens\fechar_popup.PNG" 
        self.click_specific_button_wise_1(fechar_popup)
        time.sleep(2)
        fechar_popup_2 = r"C:\Users\user\Documents\RPA_Project\imagens\fechar_popup_2.PNG" 
        self.click_specific_button_wise_1(fechar_popup_2)
        time.sleep(2)
        download = r"C:\Users\user\Documents\RPA_Project\imagens\download.PNG" 
        self.click_specific_button_wise(download)
        time.sleep(2)
        self.save_as(num_docto, id_solicitacao)
        if download_buttons_visible:
            for numero_docto, download_button in zip(numeros_docto, download_buttons):
                valor_docto = numero_docto.get_attribute('value')
                if str(valor_docto) == str(num_docto) and download_button.is_displayed():
                    download_button.click()
                    break
        # else:
            # for numero_docto, link_element in zip(numeros_docto, link_elements):
            #     valor_docto = numero_docto.get_attribute('value')
            #     if str(valor_docto) == str(num_docto) and link_element.is_displayed():
            #         link_element.click()
            #         break
            #     time.sleep(4)
            # download = r"C:\Users\user\Documents\RPA_Project\imagens\download.PNG" 
            # self.click_specific_button_wise(download)
            # time.sleep(2)
            # self.save_as(num_docto, id_solicitacao)

    def save_as(self, num_docto, id_solicitacao):
        title = "Salvar como"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de salvar como não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='win32').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Salvar como")
            return 
        janela = app[title]
        get_document = r"C:\Users\user\Documents\RPA_Project\imagens\aps.PNG" 
        self.click_specific_button_wise(get_document)
        time.sleep(2)
        file_name = janela.child_window(class_name="Edit")
        # file_name.type_keys(f"nota_{num_docto}{id_solicitacao}")
        self.type_slowly(file_name, f"nota_{num_docto}{id_solicitacao}")
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        self.fechar_aba()
    
    def fechar_aba(self):
        time.sleep(1)
        google_focus = r"C:\Users\user\Documents\RPA_Project\imagens\google_focus.PNG"
        time.sleep(1)
        self.click_specific_button_wise(google_focus)
        pyautogui.hotkey('ctrl','w')

    def type_slowly(self, element, text, delay=0.3):
        """Types the text into the element with a delay between each keystroke."""
        for character in text:
            element.type_keys(character)
            time.sleep(delay)

    def reload_page(self):
        if self.driver:
            self.driver.refresh()
            time.sleep(3)

    def get_xml(self, cnpj, num_nota, id_solicitacao):
        self.reload_page()
        if self.driver:
            self.driver.get("https://hub.sieg.com/")
            if self.driver.find_elements(By.NAME, "txtEmail") and self.driver.find_elements(By.NAME, "txtPassword"):
                campo_email = self.driver.find_element(By.NAME, "txtEmail")
                campo_email.send_keys("filipe@wisemanager.com.br")
                time.sleep(1)
                campo_senha = self.driver.find_element(By.NAME, "txtPassword")
                campo_senha.send_keys("Filipe@2023")
                time.sleep(1)
                botao_entrar = self.driver.find_element(By.NAME, "btnSubmit")
                botao_entrar.click()          
            time.sleep(2)
            # if self.driver.find_elements(By.CSS_SELECTOR, '.btn-close-alert[data-dismiss="modal"]'):      
            #     btn_close_alert = WebDriverWait(self.driver, 10).until(
            #     EC.presence_of_element_located((By.CSS_SELECTOR, '.btn-close-alert[data-dismiss="modal"]'))
            #     )
            #     btn_close_alert.click()
            botao_xml = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//span[text()=\"Xml's Baixados\"]")))
            botao_xml.click()
            time.sleep(2)
            botao_opcoes = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, "btnOptions")))
            botao_opcoes.click()
            time.sleep(2)

            campo_cnpj = WebDriverWait(self.driver, 30).until(
                EC.visibility_of_element_located((By.ID, "cnpjEmitDownload"))
            )
            campo_cnpj.send_keys(cnpj) 

            campo_numero = WebDriverWait(self.driver, 30).until(
                EC.visibility_of_element_located((By.ID, "numberXml"))
            )
            campo_numero.send_keys(num_nota)
            campo_texto = WebDriverWait(self.driver, 60).until(EC.presence_of_element_located((By.ID, "xmlKeyDownload")))
            # campo_texto.send_keys(chave_acesso)
            time.sleep(2)
            botao_pesquisar = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.XPATH, "//a[contains(@onclick, 'SearchXmlDownload')]")))
            botao_pesquisar.click()
            time.sleep(2)
            try:
                botao_detalhes = WebDriverWait(self.driver, 30).until(EC.element_to_be_clickable((By.CLASS_NAME, "btnDetailsXml")))
                botao_detalhes.click()
            except:
                self.fechar_aba()
                self.send_message_case_not_sieg(id_solicitacao, num_nota)
                database.autoriza_rpa_para_n(id_solicitacao)
            time.sleep(3)
            # valor_chave_acesso = WebDriverWait(self.driver, 20).until(
            #     EC.presence_of_element_located((By.ID, "xmlKey"))
            # ).text
            # print(valor_chave_acesso)
            # print(self.get_valor_chave_acesso())
            # botao_download = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, "btnDownloadXmlHub")))
            # botao_download.click()

    def send_message_case_not_sieg(self, id_solicitacao, numero_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Nota nao encontrada no SIEG, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}"
        
        base_url = f"https://api.telegram.org/bot{token}/sendMessage"

        for chat_id in chat_ids:
            payload = {
                'chat_id': chat_id,
                'text': success_msg
            }
            response = requests.post(base_url, data=payload)
            if response.status_code != 200:
                print(f"Failed to send success message to chat_id {chat_id}. Response: {response.content}")
            else:
                print(f"Success message sent successfully to chat_id {chat_id} on Telegram!")

    def get_valor_chave_acesso(self):
        if self.driver:
            time.sleep(2)
            valor_chave_acesso = WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "xmlKey"))
            ).text
        return valor_chave_acesso
            
    
    def download_nota(self):
        if self.driver:
            botao_download = WebDriverWait(self.driver, 60).until(EC.element_to_be_clickable((By.ID, "btnDownloadXmlHub")))
            botao_download.click()

    def rename_last_downloaded_file(self, num_nota, num_solicitacao, download_folder=None):
        if download_folder is None:
            download_folder = str(Path.home() / "Downloads")

        last_downloaded_file = max([f for f in os.scandir(download_folder)], key=lambda x: x.stat().st_mtime)

        new_name = f"xml_{num_nota}_{num_solicitacao}.xml"

        new_file_path = os.path.join(download_folder, new_name)

        os.rename(last_downloaded_file.path, new_file_path)

        print(f"Arquivo renomeado para: {new_file_path}")



    # def check_if_file_exists_in_downloads(self, num_nota, num_solicitacao, download_folder=None):
    #     if download_folder is None:
    #         download_folder = str(Path.home() / "Downloads")

    #     expected_file_name = f"xml_{num_nota}_{num_solicitacao}.xml"
    #     expected_file_path = os.path.join(download_folder, expected_file_name)

    #     # Verificar se o arquivo existe
    #     if not os.path.isfile(expected_file_path):
    #         return False

    #     try:
    #         # Tentar parsear o arquivo XML para verificar se está bem formado
    #         ET.parse(expected_file_path)
    #         return True  # O arquivo existe e está bem formado
    #     except ET.ParseError:
    #         return False  # O arquivo existe, mas não está bem formado


# teste = wise.check_if_file_exists_in_downloads('1','2')
# print(teste)
# wise.get_nf_values('109965','149495')
# wise.Anexar_AP('148624', '123', '135814', '1')
# time.sleep(2)
# wise.get_pdf_file('6','715148316')
