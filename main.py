import time
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
import database
import pyautogui
from wisemanager import Wise
import subprocess
# import pdfplumber
# import re

class NbsRpa():
    
    def __init__(self):
        pass

    def open_application(self):
        pyautogui.hotkey('win', 'm')
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title="Program Manager")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Program Manager")
            return 
        janela = app['Program Manager']
        abrir_nbs = janela.child_window(title="NBSShotcut")
        abrir_nbs.double_click_input()

    def remote_login(self):
        try:
            app = Application(backend='uia').connect(title="Acesso Remoto")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Acesso Remoto")
            return 
        janela = app['Acesso Remoto']
        conectar = janela.child_window(class_name="Button", title="Conectar")
        conectar.click_input()  

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
        password.type_keys("gerati2023")
        time.sleep(0.5)
        server.type_keys("nbs")
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
        button_image_path = r"C:\Users\user\Documents\RPA_Project\imagens\Inserir_Nota.PNG"
        self.click_specific_button(button_image_path)

    def janela_cadastro_nf(self, cpf_cnpj_value, num_nf_value, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value, numeroos, terceiro, estado, usa_rateio_centro_custo, valor_sg, rateios, rateios_aut):
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
        raio = r"C:\Users\user\Documents\RPA_Project\imagens\Raio.PNG"
        gerar = r"C:\Users\user\Documents\RPA_Project\imagens\Gerar.PNG"
        tab_faturamento = janela.child_window(control_type="TabItem", found_index=2)
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
        excluir = r"C:\Users\user\Documents\RPA_Project\imagens\Excluir.PNG"
        descricao.type_keys(contab_descricao_value)
        time.sleep(1)
        self.click_specific_button(raio)
        time.sleep(1)
        # if usa_rateio_centro_custo == 'N':
        alterar = r"C:\Users\user\Documents\RPA_Project\imagens\Alterar.PNG"
        self.click_specific_button(alterar)
        self.get_conta_contabil()
        conta_contabil = self.get_conta_contabil()
        time.sleep(2)
        self.fechar_conta_contabilizacao()
        informacao_dialog = janela.child_window(title="Informação")
        ok_button = informacao_dialog.child_window(class_name="Button")
        loop_continue = True
        while loop_continue:
            self.click_specific_button(excluir)
            descricao.click_input()
            for elem in janela.children():
                if "Informação" in elem.window_text():
                    ok_button.click()
                    loop_continue = False
                    break
        time.sleep(2)
        self.inserir_rateio(rateios, conta_contabil, valor_sg, valor_value, usa_rateio_centro_custo, rateios_aut)
        time.sleep(2)
        tab_faturamento.click_input()
        time.sleep(1)
        total_parcelas.type_keys(total_parcelas_value)
        if total_parcelas_value > 1:
            intervalo.type_keys("30")
        time.sleep(1)
        self.click_specific_button(gerar)
        time.sleep(1)
        tipo_pagamento.click_input()
        if tipo_pagamento_value in ('B','A'):
            pyautogui.typewrite('Boleto Bancario')
            pyautogui.press('tab')
        natureza_despesa.click_input()
        time.sleep(1)
        pyautogui.typewrite(natureza_financeira_value)
        time.sleep(1)
        pyautogui.press('tab')
        print(terceiro)
        if terceiro == 'S':
            self.janela_se_terceiro(numeroos)
        time.sleep(2)
        submit_button.click_input()
        if estado != 'SC':
            pyautogui.press('tab')
            pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')

    def janela_se_terceiro(self, numeroos):
        try:
            app = Application(backend='uia').connect(title="Entrada Diversas / Operação: 52-Entrada Diversas")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de nf")
            return 
        janela = app['Entrada Diversas / Operação: 52-Entrada Diversas'] 
        relacao_os = janela.child_window(control_type="TabItem", found_index=3)
        cadeado = r"C:\Users\user\Documents\RPA_Project\imagens\Cadeado.PNG"
        pesquisa_os = r"C:\Users\user\Documents\RPA_Project\imagens\Procura_Os.PNG"
        numero_os = janela.child_window(class_name="TOvcNumericField", found_index=0)
        if numeroos != '':
            relacao_os.click_input()
            time.sleep(1)
            numero_os.click_input()
            time.sleep(1)
            pyautogui.typewrite(numeroos)
            time.sleep(1)
            self.click_specific_button(pesquisa_os)
        else:
            relacao_os.click_input()
            time.sleep(2)
            self.click_specific_button(cadeado)
            time.sleep(0.2)
            pyautogui.press('tab')
            time.sleep(0.2)
            pyautogui.press('enter')

    def janela_valores(self):
        try:
            app = Application(backend='uia').connect(title="Entrada Diversas / Operação: 52-Nota de despesas Diversas")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de nota de despesas")
            return 
        janela = app['Entrada Diversas / Operação: 52-Nota de despesas Diversas']
        tab_dados = janela.child_window(control_type="TabItem", found_index=1)
        tab_dados.click_input()
        pyautogui.press("enter")
    
    def get_controle_value(self):
        try:
            app = Application(backend='uia').connect(title="Entrada Diversas / Operação: 52-Nota de despesas Diversas")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de nota de despesas")
            return 
        janela = app['Entrada Diversas / Operação: 52-Nota de despesas Diversas']
        controle = janela.child_window(class_name="TDBEdit", found_index=2)
        properties = controle.legacy_properties()
        return properties['Value']
    
    def janela_imprimir_nota(self):
        try:
            app = Application(backend='uia').connect(title="Ficha de Controle de Pagamento")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de imprimir nota")
            return 
        janela = app['Ficha de Controle de Pagamento']
        # janela.wait('visible', timeout=120)
        imprimir = r"C:\Users\user\Documents\RPA_Project\imagens\Imprimir.PNG"
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

    def save_as(self, num_docto):
        try:
            app = Application(backend='uia').connect(title="Salvar como")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Salvar como")
            return 
        janela = app['Salvar como']
        # janela.wait('visible', timeout=120)
        # Declare Variables
        get_document = janela.child_window(title="APs (fixo)")
        file_name = janela.child_window(class_name="Edit")
        file_type = janela.child_window(class_name="AppControlHost", found_index=1)
        # Set Comands
        get_document.click_input()
        time.sleep(2)
        file_name.type_keys(f"AP_{num_docto}")
        pyautogui.press('tab')
        file_type.click_input()
        for _ in range(2):
            pyautogui.press('down')
            time.sleep(0.5)
        pyautogui.press('enter')
        for _ in range(2):
            pyautogui.press('enter')
            time.sleep(1)
        time.sleep(4)
        pyautogui.hotkey('ctrl', 'w')

    def click_specific_button(self, button_image_path, confidence_level=0.7):
        button_location = pyautogui.locateOnScreen(button_image_path, confidence=confidence_level)
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

    def close_aplications_half(self):
        janela_anterior_nbs_entrada = r"C:\Users\user\Documents\RPA_Project\imagens\Janela.PNG"
        time.sleep(2)
        self.click_specific_button(janela_anterior_nbs_entrada)

    def close_aplications_end(self):
        janela_anterior_sistema_financeiro = r"C:\Users\user\Documents\RPA_Project\imagens\Sair.PNG"
        self.click_specific_button(janela_anterior_sistema_financeiro)
        time.sleep(2)
        subprocess.run(['taskkill', '/F', '/IM', 'NBSShortCut.exe'])
    
    def click_on_cancel(self):
        cancel_pag = r"C:\Users\user\Documents\RPA_Project\imagens\Cancelar.PNG"
        self.click_specific_button(cancel_pag)

    # def get_controle_pdf(self, num_docto):
    #     caminho_pdf = rf"C:\Users\user\Documents\AP_{num_docto}.pdf"
    #     with pdfplumber.open(caminho_pdf) as pdf:
    #         primeira_pagina = pdf.pages[0]
    #         texto = primeira_pagina.extract_text() 
    #     palavras = texto.split()
    #     formato_especifico = re.compile(r'^\d{1,2}-\d+-\d$')
    #     for palavra in palavras:
    #         if formato_especifico.match(palavra):
    #             partes = palavra.split('-')
    #             return partes[1]  
    #     return None
    
    def back_to_nbs(self):
        try:
            app = Application(backend='uia').connect(title="Barra de Tarefas")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Barra de Tarefas")
            return 
        janela = app['Barra de Tarefas']
        minimizar_google = janela.child_window(title="Google Chrome - 1 executando o windows")
        minimizar_google.click_input()
        
    def fechar_conta_contabilizacao(self):
        try:
            app = Application(backend='uia').connect(title="Alterar Conta de Contabilização")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de alterar conta de contabilizacao")
            return 
        janela = app['Alterar Conta de Contabilização']
        fechar = janela.child_window(class_name='TBitBtn', found_index=0)
        fechar.click_input()
   
    def inserir_rateio(self, rateios, conta_contabil, valor_sg, vlr_nf, usa_rateio_centro_custo, rateios_aut):
        incluir = r"C:\Users\user\Documents\RPA_Project\imagens\Incluir.PNG"
        self.click_specific_button(incluir)
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title="Incluir Conta de Contabilização")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de alterar conta de contabilizacao")
            return 
        janela = app['Incluir Conta de Contabilização']
        contacontabil = janela.child_window(class_name="TNBSContaContab")
        centro_custo = janela.child_window(found_index=2)
        historico_padrao = janela.child_window(found_index=3)
        confirmar = janela.child_window(class_name='TBitBtn', found_index=1)
        valor_rateio = janela.child_window(class_name="TOvcPictureField", found_index=2)
        if usa_rateio_centro_custo == 'N':
            for rateio in rateios:
                self.click_specific_button(incluir)
                contacontabil.type_keys(conta_contabil)
                pyautogui.press("tab")
                centro_custo.type_keys(rateio[0]) 
                historico_padrao.click_input()
                pyautogui.typewrite("85")
                vlr_rateio = self.get_valor_rateio_manual(rateio[1], valor_sg, vlr_nf)
                valor_rateio.type_keys(vlr_rateio)
                print(vlr_rateio)
                confirmar.click_input()
                
        elif usa_rateio_centro_custo == 'S':
            for rateio_aut in rateios_aut:
                self.click_specific_button(incluir)
                contacontabil.type_keys(conta_contabil)
                pyautogui.press("tab")
                centro_custo.type_keys(rateio_aut[0]) 
                historico_padrao.click_input()
                pyautogui.typewrite("85")
                vlr_rateio = self.get_valor_rateio_aut(vlr_nf, rateio_aut[1])
                valor_rateio.type_keys(vlr_rateio)
                print(vlr_rateio)
                confirmar.click_input()

    def get_valor_rateio_manual(self, valor_centro_custo, total_sg, valor_nf):
        valor_nf = valor_nf.replace(',', '.')
        valor_nf = float(valor_nf) if valor_nf else 0.0  
        proporcao_centro_custo = valor_centro_custo / total_sg
        valor_rateio = proporcao_centro_custo * valor_nf
        valor_rateio_str = "{:.2f}".format(valor_rateio)
        valor_rateio_str = valor_rateio_str.lstrip('0')
        return valor_rateio_str.replace('.', ',')

    def get_valor_rateio_aut(self, valor_nf, porcentagem):
        valor_nf = valor_nf.replace(',', '.')
        valor_nf = float(valor_nf) if valor_nf else 0.0  
        porcentagem_decimal = float(porcentagem) / 100
        valor_rateio = porcentagem_decimal * valor_nf
        valor_rateio_str = "{:.2f}".format(valor_rateio)
        valor_rateio_str = valor_rateio_str.lstrip('0')
        return valor_rateio_str.replace('.', ',')

    def get_conta_contabil(self):
        try:
            app = Application(backend='uia').connect(title="Alterar Conta de Contabilização")
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de alterar conta de contabilizacao")
            return 
        janela = app['Alterar Conta de Contabilização']
        time.sleep(2)
        conta_contabil = janela.child_window(class_name="TNBSContaContab")
        conta_contabil_value = conta_contabil.legacy_properties()
        print(conta_contabil_value['Value'])
        return conta_contabil_value['Value']
    
    def funcao_main(self):
        registros = database.consultar_dados_cadastro()
        empresa_anterior = None
        for row in registros:
            empresa_atual = row[2]
            if empresa_atual != empresa_anterior:
                self.close_aplications_end()
                time.sleep(3)
                self.open_application()
                time.sleep(2)
                self.remote_login()
                time.sleep(5)
                self.login()
                time.sleep(7)
                self.janela_empresa_filial(row[2], row[3])
            time.sleep(5)
            empresa_anterior = empresa_atual
            print(empresa_anterior, empresa_atual)
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
            usa_rateio_centro_custo = row[14] #
            valor_sg = row[15] # 
            id_rateiocc = row[16]
            rateios_aut = database.consultar_rateio_aut(id_rateiocc)
            notas_fiscais = database.consultar_nota_fiscal(id_solicitacao)
            if notas_fiscais:
                numerodocto = notas_fiscais[0][2]
                serie_value = notas_fiscais[0][3]
                data_emissao_value = notas_fiscais[0][1] 
                tipo_docto_value = notas_fiscais[0][5]
                valor_value = notas_fiscais[0][0]
            rateios = database.consultar_rateio(id_solicitacao)
            numeroos = row[11]
            terceiro = row[12]
            estado = row[13]
            self.janela_cadastro_nf(cnpj, numerodocto, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value, numeroos, terceiro, estado, usa_rateio_centro_custo, valor_sg, rateios, rateios_aut)
            time.sleep(10)
            self.janela_imprimir_nota()
            time.sleep(5)
            self.janela_secundario_imprimir_nota()
            time.sleep(5)
            self.extract_pdf()
            time.sleep(3)
            self.save_as(numerodocto)
            time.sleep(5)
            self.close_extract_pdf_window()
            time.sleep(2)
            self.click_on_cancel()
            time.sleep(5)
            self.janela_valores()
            time.sleep(2)
            num_controle = self.get_controle_value()
            time.sleep(2)
            print(num_controle)
            time.sleep(5)
            wise_instance = Wise()
            time.sleep(3)
            wise_instance.Anexar_AP(id_solicitacao, num_controle, numerodocto)
            time.sleep(3)
            wise_instance.get_pdf_file(numerodocto)
            time.sleep(4)
            wise_instance.confirm()
            time.sleep(3)
            database.atualizar_anexosolicitacaogasto(numerodocto)
            time.sleep(2)
            self.back_to_nbs()
            time.sleep(2)
            self.close_aplications_half()
            time.sleep(3)

rpa = NbsRpa()
rpa.funcao_main()
