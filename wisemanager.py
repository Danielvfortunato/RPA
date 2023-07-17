from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
import pyautogui


class Wise():
    def __init__(self):
        self.driver = None

    def login(self, id_solicitacao, num_controle):
        # Configurar o driver
        chrome_options = Options()
        chrome_options.add_experimental_option("detach", True)
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.maximize_window() 
        self.driver.get("https://demow.wisemanager.com.br/WiseManagerBI/")

        # Localizar os campos de login
        username_field = self.driver.find_element(By.NAME, "j_username")
        password_field = self.driver.find_element(By.NAME, "j_password")
        # Preencher e enviar os campos de login
        username_field.send_keys("Daniel")
        time.sleep(1)
        password_field.send_keys("Da131800!")
        time.sleep(0.5)
        password_field.send_keys(Keys.RETURN)

        #Preencher registro
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
        time.sleep(5)
        open_detail = self.driver.find_element(By.XPATH, "//button[contains(@ng-click, 'abreDetalheSolicitacao')]")
        open_detail.click()
        time.sleep(2)
        controle = self.driver.find_element(By.NAME, "numeroControle")
        controle.send_keys(num_controle)
        time.sleep(3)
        file_input = self.driver.find_element(By.XPATH, "//input[@class='file-input']")
        self.driver.execute_script("arguments[0].click();", file_input)
    

    def get_pdf_file(self, id_solicitacao):
        try:
            app = Application(backend="uia").connect(title="WiseManager - Google Chrome")
        except ElementNotFoundError:
            print("Falha ao conectar com página pegar pdf")
            return 
        janela = app['WiseManager - Google Chrome']

        get_document = janela.child_window(title="Documentos", found_index=0)
        file_name = janela.child_window(class_name="Edit")
        
        get_document.click_input()
        file_name.type_keys(f"AP_{id_solicitacao}.PDF")
        file_name.type_keys("{ENTER}")

    def confirm(self):
        url_esperada = "https://demow.wisemanager.com.br/WiseManagerBI/#/financeiro/efetivacaoSolicitacaoGastoNew"

        if self.driver.current_url == url_esperada:
            botao_confirmar = self.driver.find_element(By.XPATH, "//button[contains(@title, 'Confirma Lançamento no ERP')]")
            botao_confirmar.click()
            time.sleep(4)
            confirmar = r"C:\Users\user\Pictures\Confirmar.PNG"
            self.click_specific_button_wise(confirmar)
            time.sleep(2)
            pyautogui.hotkey('ctrl', 'w')
            # confirmar = self.driver.find_element(By.XPATH, "//button[contains(@ng-click, 'efetivaSolicitacaoGasto')]")
            # confirmar.click()
        else:
            print("Página nao encontrada")

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


