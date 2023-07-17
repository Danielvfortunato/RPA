import time
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
from pywinauto.keyboard import SendKeys
import database
import pyautogui
from wisemanager import Wise
import pdfplumber
# import tratamento_dados



class NbsRpa():

    def __init__(self):
        pass

    def open_application(self):
        try:
            app = Application().start(r"C:\NBS\SisFin.exe")
        except ElementNotFoundError:
            print("Falha ao abrir aplicativo")
            return    

    def login(self):
        try:
            app = Application(backend="uia").connect(title="NBS - Login")
        except ElementNotFoundError:
            print("Falha ao conectar com página de login")
            return 
        janela = app.window(title="NBS - Login")
        janela.wait('visible', timeout=120)
        # Declare Variables
        user = janela.child_window(class_name="TOvcPictureField", found_index=1)
        password = janela.child_window(class_name="TOvcPictureField", found_index=0)
        server = janela.child_window(class_name="TButtonedEdit")
        submit = janela.child_window(class_name="TfcImageBtn", found_index=0)
        # Set Comands
        user.type_keys("NBS")
        time.sleep(0.5)
        password.type_keys("new")
        time.sleep(0.5)
        server.type_keys("geracaobkp")
        time.sleep(0.5)
        submit.click_input()
        
    def initial_page(self):
        try:
            app = Application(backend="uia").connect(title="NBS ShortCut")
        except ElementNotFoundError:
            print("Falha ao conectar ao aplicativo")
            return
        janela = app.window(title="NBS ShortCut")
        # janela.wait('visible', timeout=120)
        # Declare Variables
        apply_sisfin = janela.child_window(class_name="TEdit")
        # Set Comands
        apply_sisfin.set_text("SisFin")
        apply_sisfin.type_keys("{ENTER}")

    def janela_empresa_filial(self, empresa_name_value, cod_matriz_value):
        try:
            app = Application(backend="uia").connect(title='Empresa/Filial')
        except ElementNotFoundError:
            print("Erro ao conectar com janela empresa filial")
            return
        janela = app['Empresa/Filial']
        # janela.wait('visible', timeout=120)
        # Declare Variables
        empresa = janela.child_window(class_name='TDBLookupComboBox', found_index=1)
        filial = janela.child_window(class_name="TDBLookupComboBox", found_index=0)
        confirma = janela.child_window(class_name="TBitBtn", found_index=1)
        # Set Comands
        empresa.type_keys(cod_matriz_value)
        time.sleep(1)
        filial.click_input()
        time.sleep(1)
        pyautogui.typewrite(empresa_name_value)
        filial.type_keys('{ENTER}')
        confirma.click_input()
        
    def access_contas_a_pagar(self):
        try:
            app = Application(backend="win32").connect(title='Sistema Financeiro - SISFIN')
        except ElementNotFoundError:
            print("Erro ao conectar com janela Sistema Financeiro")
            return
        janela = app['Sistema Financeiro - SISFIN']
        # janela.wait('visible', timeout=120)
        # Set Comands
        janela.set_focus()
        pyautogui.hotkey('alt', 'p')
        time.sleep(1)
        pyautogui.press('n')

    def janela_entrada(self):
        try:
            app = Application(backend="win32").connect(title="Entradas")
        except ElementNotFoundError:
            print("Erro ao conectar com janela de entradas")
            return 
        janela = app['Entradas']
        # janela.wait('visible', timeout=120)
        janela.set_focus()
        time.sleep(2)
        button_image_path = r"C:\Users\user\Pictures\Captura.PNG"
        self.click_specific_button(button_image_path)

    def janela_cadastro_nf(self, cpf_cnpj_value, num_nf_value, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value):
        try:
            app = Application(backend='uia').connect(title="Entrada Diversas / Operação: 52-Entrada Diversas")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de nf")
            return 
        janela = app['Entrada Diversas / Operação: 52-Entrada Diversas']
        # janela.wait('visible', timeout=120)
        # Declare Variables
        cpf_cnpj = janela.child_window(class_name="TCPF_CGC")
        esp = janela.child_window(class_name="TOvcDbPictureField", found_index=2)
        num_nf = janela.child_window(class_name="TOvcDbPictureField", found_index=6)
        sr = janela.child_window(class_name="TOvcDbPictureField", found_index=5)
        data_emissao = janela.child_window(class_name="TOvcDbPictureField", found_index=4)
        vlr_nf = janela.child_window(class_name="TOvcDbPictureField", found_index=20)
        tab_contab = janela.child_window(control_type="TabItem", found_index=2)
        descricao = janela.child_window(class_name="TwwDBLookupCombo")
        # contab_cod = janela.child_window(class_name="TOvcNumericField", found_index=0)
        raio = r"C:\Users\user\Pictures\Raio.PNG"
        gerar = r"C:\Users\user\Pictures\Gerar_2.PNG"
        tab_faturamento =janela.child_window(found_index=41)
        total_parcelas = janela.child_window(class_name="TOvcNumericField", found_index=0)
        intervalo = janela.child_window(class_name="TOvcNumericField", found_index=1)
        # gerar
        tipo_pagamento = janela.child_window(class_name="TwwDBLookupCombo", found_index=1)
        natureza_despesa = janela.child_window(class_name="TwwDBLookupCombo", found_index=0)
        submit_button = janela.child_window(class_name="TPanel", found_index=0)
        # Set Comands
        cpf_cnpj.type_keys(cpf_cnpj_value)
        cpf_cnpj.type_keys("{ENTER}")
        esp.click_input()
        esp.type_keys(tipo_docto_value)
        num_nf.type_keys(num_nf_value)
        sr.type_keys(serie_value)
        # data_emissao.type_keys(data_emissao_value) Adicionar depois
        vlr_nf.type_keys(valor_value)
        tab_contab.click_input()
        time.sleep(0.1) 
        pyautogui.press('right')
        descricao.type_keys(contab_descricao_value)
        time.sleep(1)
        self.click_specific_button(raio)
        tab_faturamento.click_input()
        time.sleep(1)
        total_parcelas.type_keys(total_parcelas_value)
        if total_parcelas_value > 1:
            intervalo.type_keys("30")
        time.sleep(1)
        self.click_specific_button(gerar)
        time.sleep(1)
        tipo_pagamento.click_input()
        pyautogui.typewrite(tipo_pagamento_value)
        pyautogui.press('tab')
        # tipo_pagamento.click_input()
        # for _ in range(2):
        #     pyautogui.press('down')
        #     time.sleep(0.1)
        # pyautogui.press('enter')
        # natureza_despesa.click_input()
        # for _ in range(5):
        #     pyautogui.press('down')
        # pyautogui.press('enter')
        natureza_despesa.click_input()
        time.sleep(1)
        pyautogui.typewrite(natureza_financeira_value)
        time.sleep(1)
        pyautogui.press('tab')
        submit_button.click_input()
        time.sleep(1)
        pyautogui.press('enter')

    def janela_imprimir_nota(self):
        try:
            app = Application(backend='uia').connect(title="Ficha de Controle de Pagamento")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de imprimir nota")
            return 
        janela = app['Ficha de Controle de Pagamento']
        # janela.wait('visible', timeout=120)
        imprimir = r"C:\Users\user\Pictures\Imprimir2.PNG"
        self.click_specific_button(imprimir)

    def janela_secundario_imprimir_nota(self):
        try:
            app = Application(backend='uia').connect(title="Report Destination")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela secundaria de imprimir nota")
            return 
        janela = app['Report Destination']
        # janela.wait('visible', timeout=120)
        # Declare Variables
        v_print = janela.child_window(found_index=12)
        # Set Comands
        v_print.click_input()

    def extract_pdf(self):
        try:
            app = Application(backend='uia').connect(title="Ace Viewer")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Ace Viewer")
            return 
        janela = app['Ace Viewer']
        # janela.wait('visible', timeout=120)
        # janela.set_focus()
        pyautogui.hotkey('alt', 'f')
        time.sleep(1)
        pyautogui.press('s')

    def save_as(self, solicitacao_gasto):
        try:
            app = Application(backend='uia').connect(title="Salvar como")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Salvar como")
            return 
        janela = app['Salvar como']
        # janela.wait('visible', timeout=120)
        # Declare Variables
        file_name = janela.child_window(class_name="Edit")
        file_type = janela.child_window(class_name="AppControlHost", found_index=1)
        # Set Comands
        file_name.type_keys(f"AP_{solicitacao_gasto}")
        pyautogui.press('tab')
        file_type.click_input()
        for _ in range(2):
            pyautogui.press('down')
            time.sleep(0.5)
        pyautogui.press('enter')
        for _ in range(2):
            pyautogui.press('enter')
            time.sleep(1)
        time.sleep(1)
        pyautogui.hotkey('ctrl', 'w')

    def click_specific_button(self, button_image_path):
        button_location = pyautogui.locateOnScreen(button_image_path)
        if button_location:
            button_x, button_y, button_width, button_height = button_location
            button_center_x = button_x + button_width // 2
            button_center_y = button_y + button_height // 2
            pyautogui.click(button_center_x, button_center_y)
            print("Botão clicado!")
        else:
            print("Botão não encontrado.")

    
    def close_extract_pdf_window(self):
        try:
            app = Application(backend='uia').connect(title="Ace Viewer")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Ace Viewer")
            return 
        janela = app['Ace Viewer']
        # janela.wait('visible', timeout=120)
        janela.set_focus()
        fechar = janela.child_window(title="Fechar")
        fechar.click_input()

    def close_aplications_by_cancel_icon(self):
        cancel_pag = r"C:\Users\user\Pictures\Cancel_Pag.PNG"
        janela_anterior_nbs_entrada = r"C:\Users\user\Pictures\Janela_Anterior_NBS.PNG"
        janela_anterior_sistema_financeiro = r"C:\Users\user\Pictures\Janela_Anterior_Sis_Financ.PNG"
        self.click_specific_button(cancel_pag)
        time.sleep(2)
        self.click_specific_button(janela_anterior_nbs_entrada)
        time.sleep(2)
        self.click_specific_button(janela_anterior_sistema_financeiro)

    def get_controle_pdf(self, solicitacao_gasto):
        caminho_pdf = rf"C:\Users\user\Documents\AP_{solicitacao_gasto}.pdf"
        with pdfplumber.open(caminho_pdf) as pdf:
            primeira_pagina = pdf.pages[0]  
            texto = primeira_pagina.extract_text() 
        palavras = texto.split()
        indices_controle = [i for i, palavra in enumerate(palavras) if palavra == 'Controle']
        if len(indices_controle) < 2:
            print("Não foram encontradas duas ocorrências de 'Controle'.")
        else:
            valor_controle = palavras[indices_controle[1] + 1]
            valor_controle = valor_controle.replace("-", "")
        return valor_controle
    

    def funcao_main(self):
        registros = database.consultar_dados_cadastro()
        for row in registros:
            self.open_application()
            time.sleep(3)
            self.login()
            time.sleep(5)
            self.janela_empresa_filial(row[2], row[3])
            time.sleep(5)
            self.access_contas_a_pagar()
            time.sleep(5)
            self.janela_entrada() 
            time.sleep(5)
            id_solicitacao = row[0]
            cnpj = row[1]
            contab_descricao_value = row[6]
            total_parcelas_value = row[9]
            tipo_pagamento_value = row[5]
            natureza_financeira_value = row[8]
            notas_fiscais = database.consultar_nota_fiscal(id_solicitacao)
            if notas_fiscais:
                numerodocto = notas_fiscais[0][2]
                serie_value = notas_fiscais[0][3]
                data_emissao_value = notas_fiscais[0][1] 
                tipo_docto_value = notas_fiscais[0][5]
                valor_value = notas_fiscais[0][0]
            self.janela_cadastro_nf(cnpj, numerodocto, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value)
            time.sleep(7)
            self.janela_imprimir_nota()
            time.sleep(5)
            self.janela_secundario_imprimir_nota()
            time.sleep(5)
            self.extract_pdf()
            time.sleep(5)
            self.save_as(id_solicitacao)
            time.sleep(3)
            wise_instance = Wise()
            num_controle = self.get_controle_pdf(id_solicitacao)
            time.sleep(3)
            wise_instance.login(id_solicitacao, num_controle)
            time.sleep(3)
            wise_instance.get_pdf_file(id_solicitacao)
            time.sleep(4)
            wise_instance.confirm()
            # self.close_extract_pdf_window()
            # time.sleep(2)
            # self.close_aplications_by_cancel_icon()

rpa = NbsRpa()
rpa.funcao_main()
