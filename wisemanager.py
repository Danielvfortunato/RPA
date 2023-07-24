from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
import psutil
import subprocess
import time
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
import pyautogui
from selenium.webdriver.support.select import Select

class Wise():
    def __init__(self):
        self.driver = None

    def start_chrome_debugger(self):
        for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
            if proc.info['name'] == 'chrome.exe' and '--remote-debugging-port=9222' in proc.info['cmdline']:
                time.sleep(1)
                self.open_google()
                print("Chrome com depurador remoto já está em execução.")
                return
        try:
            subprocess.Popen([
                "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
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
        abrir_google.click_input()

    def init_instance_chrome(self):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
        if not self.driver:
            chrome_driver = "C:\chromedriver.exe"
            self.driver = webdriver.Chrome(executable_path=chrome_driver, options=chrome_options)
        return self.driver

    def login(self):
            username_field = self.driver.find_element(By.NAME, "j_username")
            password_field = self.driver.find_element(By.NAME, "j_password")
            username_field.send_keys('Daniel')
            password_field.send_keys('Da131800!')
            password_field.submit()

    def Anexar_AP(self, id_solicitacao, num_controle, num_docto):
        self.start_chrome_debugger()
        self.init_instance_chrome()
        time.sleep(2)
        self.driver.get("https://demow.wisemanager.com.br/WiseManagerBI/#/financeiro/efetivacaoSolicitacaoGastoNew")
        if self.driver.find_elements(By.NAME, "j_username") and self.driver.find_elements(By.NAME, "j_password"):
            self.login()
            time.sleep(20)
        select_element = self.driver.find_element(By.CSS_SELECTOR, 'select.form-control.input-sm[ng-model="filtro.campo"]')
        select = Select(select_element)
        select.select_by_value('id')
        time.sleep(3)
        input_element = self.driver.find_element(By.CSS_SELECTOR, 'input.form-control.input-sm.filter-picker-round[ng-model="filtro.texto"]')
        input_element.send_keys(id_solicitacao)
        time.sleep(3)
        search = self.driver.find_element(By.CSS_SELECTOR, 'button.btn.btn-circle.btn-bordered.thick.btn-primary.imt-5')
        search.click()
        time.sleep(10)
        open_detail = self.driver.find_element(By.XPATH, "//button[contains(@ng-click, 'abreDetalheSolicitacao')]")
        open_detail.click()
        time.sleep(5)
        numeros_docto = self.driver.find_elements(By.NAME, "numeroDocto")
        controles = self.driver.find_elements(By.NAME, "numeroControle")
        for numero_docto, controle in zip(numeros_docto, controles):
            valor_docto = numero_docto.get_attribute('value')
            if valor_docto == num_docto:
                controle.send_keys(num_controle)
        # controle = self.driver.find_element(By.NAME, "numeroControle")
        # controle.send_keys(num_controle)
        time.sleep(3)
        file_input = self.driver.find_element(By.XPATH, "//input[@class='file-input']")
        self.driver.execute_script("arguments[0].click();", file_input)

    def get_pdf_file(self, numero_docto):
        try:
            app = Application(backend="uia").connect(title="WiseManager - Google Chrome")
        except ElementNotFoundError:
            print("Falha ao conectar com página pegar pdf")
            return 
        janela = app['WiseManager - Google Chrome']
        get_document = janela.child_window(title="Documentos", found_index=0)
        file_name = janela.child_window(class_name="Edit")
        get_document.click_input()
        file_name.type_keys(f"AP_{numero_docto}.PDF")
        file_name.type_keys("{ENTER}")

    def confirm(self):
        url_esperada = "https://demow.wisemanager.com.br/WiseManagerBI/#/financeiro/efetivacaoSolicitacaoGastoNew"
        if self.driver.current_url == url_esperada:
            controles = self.driver.find_elements(By.NAME, "numeroControle")
            todos_preenchidos = all(controle.get_attribute('value') != "" for controle in controles if controle.is_displayed())
            if todos_preenchidos:
                botao_confirmar = self.driver.find_element(By.XPATH, "//button[contains(@title, 'Confirma Lançamento no ERP')]")
                botao_confirmar.click()
                confirmar = r"C:\Users\user\Pictures\ConfirmarMac2.PNG"
                time.sleep(4)
                self.click_specific_button_wise(confirmar)
                time.sleep(3)
            else:
                print("Nem todos os controles estão preenchidos. Não foi possível confirmar.")
            
    def click_specific_button_wise(self, button_image_path):
        button_location = pyautogui.locateOnScreen(button_image_path)
        if button_location:
            button_x, button_y, button_width, button_height = button_location
            button_center_x = button_x + button_width // 2
            button_center_y = button_y + button_height // 2
            pyautogui.click(button_center_x, button_center_y)
            print("Botão clicado!")
        else:
            print("Botão não encontrado.")

# wise = Wise()
# wise.Anexar_AP('135167', '0000', '10783')
# wise.get_pdf_file('136042')
# wise.confirm()