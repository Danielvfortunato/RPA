import time
from pywinauto import Application
from pywinauto.findwindows import ElementNotFoundError
import database
import pyautogui
from wisemanager import Wise
import pygetwindow as gw
import re
from pytesseract import pytesseract
from pdf2image import convert_from_path
import datetime
import requests 
import pdfplumber
import traceback
import xml.etree.ElementTree as ET
import math
import send_mail

class NbsRpa():
    
    def __init__(self):
        pass
    
    def open_application(self):
        try:
            app = Application().start(r"C:\NBS\SisFin.exe")
        except ElementNotFoundError:
            print("Falha ao abrir aplicativo")
            return  

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

    def login(self):
        title = "NBS - Login"
        if not self.esperar_janela_visivel(title, timeout=120):
            print("Falha: Janela de login não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend="uia").connect(title=title)
            janela = app.window(title=title)
        except ElementNotFoundError:
            print("Falha ao conectar com página de login")
            return 
        # Declare Variables
        user = janela.child_window(class_name="TOvcPictureField", found_index=1)
        password = janela.child_window(class_name="TOvcPictureField", found_index=0)
        server = janela.child_window(class_name="TButtonedEdit")
        submit = janela.child_window(class_name="TfcImageBtn", found_index=0)
        # Set Comands
        try:
            self.wait_until_interactive(user)
        except TimeoutError as e:
            print(str(e))
            return
        user.type_keys("NBS")
        time.sleep(2)
        try:
            self.wait_until_interactive(password)
        except TimeoutError as e:
            print(str(e))
            return
        password.type_keys("gerati2023")
        time.sleep(2)
        try:
            self.wait_until_interactive(server)
        except TimeoutError as e:
            print(str(e))
            return
        server.type_keys("nbs")
        time.sleep(2)
        try:
            self.wait_until_interactive(submit)
        except TimeoutError as e:
            print(str(e))
            return
        submit.click_input()
        
    # def initial_page(self):
    #     try:
    #         app = Application(backend="uia").connect(title="NBS ShortCut")
    #     except ElementNotFoundError:
    #         print("Falha ao conectar ao aplicativo")
    #         return
    #     janela = app.window(title="NBS ShortCut")
    #     # janela.wait('visible', timeout=120)
    #     # Declare Variables
    #     apply_sisfin = janela.child_window(class_name="TEdit")
    #     # Set Comands
    #     apply_sisfin.set_text("SisFin")
    #     apply_sisfin.type_keys("{ENTER}")

    def janela_empresa_filial(self, empresa_name_value, cod_matriz_value):
        title = "Empresa/Filial"
        if not self.esperar_janela_visivel(title, timeout=120):
            print("Falha: Janela de empresa/filial não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend="uia").connect(title=title)
        except ElementNotFoundError:
            print("Erro ao conectar com janela empresa filial")
            return
        janela = app[title]
        # Declare Variables
        empresa = janela.child_window(class_name='TDBLookupComboBox', found_index=1)
        filial = janela.child_window(class_name="TDBLookupComboBox", found_index=0)
        confirma = janela.child_window(class_name="TBitBtn", found_index=1)
        # Set Comands
        try:
            self.wait_until_interactive(empresa)
        except TimeoutError as e:
            print(str(e))
            return
        empresa.type_keys(cod_matriz_value)
        try:
            self.wait_until_interactive(filial)
        except TimeoutError as e:
            print(str(e))
            return
        filial.click_input()
        time.sleep(1)
        pyautogui.typewrite(empresa_name_value)
        time.sleep(1)
        filial.type_keys('{ENTER}')
        try:
            self.wait_until_interactive(confirma)
        except TimeoutError as e:
            print(str(e))
            return
        confirma.click_input()
        
    def access_contas_a_pagar(self):
        title = "Sistema Financeiro - SISFIN"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de sistema financeiro sisfin não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend="win32").connect(title=title)
        except ElementNotFoundError:
            print("Erro ao conectar com janela Sistema Financeiro")
            return
        janela = app[title]
        # janela.wait('visible', timeout=120)
        # Set Comands
        janela.set_focus()
        pyautogui.hotkey('alt', 'p')
        time.sleep(1)
        pyautogui.press('n')

    def janela_entrada(self):
        title = "Entradas"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de entradas não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend="win32").connect(title=title)
        except ElementNotFoundError:
            print("Erro ao conectar com janela de entradas")
            return 
        janela = app[title]
        janela.set_focus()
        # time.sleep(2)
        # button_image_path = r"C:\Users\user\Documents\RPA_Project\imagens\Inserir_Nota.PNG"
        # self.click_specific_button(button_image_path)
        time.sleep(1)
        button_image_path = r"C:\Users\user\Documents\RPA_Project\imagens\nfe.PNG"
        time.sleep(1)
        self.click_specific_button(button_image_path)

    def importar_xml(self):
        title = "Arquivos NFe: Interface de compra"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de entradas não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend="win32").connect(title=title)
        except ElementNotFoundError:
            print("Erro ao conectar com janela de arquivos nfe")
            return 
        janela = app[title]
        janela.set_focus()
        xml = r"C:\Users\user\Documents\RPA_Project\imagens\xml.PNG"
        time.sleep(2)
        self.click_specific_button(xml)

    def abrir_xml(self, chave_acesso):
        title = "Abrir"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de entradas não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend="uia").connect(title=title)
        except ElementNotFoundError:
            print("Erro ao conectar com janela de abrir")
            return 
        janela = app[title]
        try:
            get_document = janela.child_window(title="Downloads (fixo)")
            get_document.click_input()  
        except:
            pass
        time.sleep(2)
        file_name = janela.child_window(class_name="Edit")
        file_name.type_keys(chave_acesso)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(3)
        pyautogui.typewrite('Hyundai_Diversos')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        aceitar = r"C:\Users\user\Documents\RPA_Project\imagens\aceitar.PNG"
        self.click_specific_button(aceitar)

    def janela_cadastro_nf(self, cpf_cnpj_value, num_nf_value, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value, numeroos, terceiro, estado, usa_rateio_centro_custo, valor_sg, rateios, rateios_aut, inss, irff, piscofinscsl, iss, vencimento_value, obs, codigo_contab, boletos, num_parcelas, id_solicitacao, chave_acesso):
        title = "Entrada Diversas / Operação: 52-Entrada Diversas"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de cadastro_nf não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de nf")
            return 
        janela = app[title]
        print(vencimento_value)
        cpf_cnpj = janela.child_window(class_name="TCPF_CGC")
        esp = janela.child_window(class_name="TOvcDbPictureField", found_index=2)
        num_nf = janela.child_window(class_name="TOvcDbPictureField", found_index=6)
        observacao = janela.child_window(class_name='TOvcDbPictureField', found_index=22)
        sr = janela.child_window(class_name="TOvcDbPictureField", found_index=5)
        data_emissao = janela.child_window(class_name="TOvcDbPictureField", found_index=4)
        vlr_nf = janela.child_window(class_name="TOvcDbPictureField", found_index=20)
        tab_contab = janela.child_window(title=' Contabilização ', control_type="TabItem")
        descricao = janela.child_window(class_name="TwwDBLookupCombo")
        raio = r"C:\Users\user\Documents\RPA_Project\imagens\Raio.PNG"
        gerar = r"C:\Users\user\Documents\RPA_Project\imagens\Gerar.PNG"
        tab_faturamento = janela.child_window(title='   Faturamento   ', control_type="TabItem")
        total_parcelas = janela.child_window(class_name="TOvcNumericField", found_index=0)
        intervalo = janela.child_window(class_name="TOvcNumericField", found_index=1)
        vencimento = janela.child_window(class_name="TOvcDbPictureField", found_index=4)
        tipo_pagamento = janela.child_window(class_name="TwwDBLookupCombo", found_index=1)
        natureza_despesa = janela.child_window(class_name="TwwDBLookupCombo", found_index=0)
        submit_button = janela.child_window(class_name="TPanel", found_index=0)
        barra_boleto = r"C:\Users\user\Documents\RPA_Project\imagens\Barra_Boleto.PNG"
        aba_cfop = janela.child_window(title='CFOP´s', control_type='TabItem')
        # Set Comands
        # cpf_cnpj.type_keys(cpf_cnpj_value)
        # cpf_cnpj.type_keys("{ENTER}")
        # esp.click_input()
        # esp.type_keys(tipo_docto_value)
        # num_nf.type_keys(num_nf_value)
        # serie = self.get_serie(num_nf_value, id_solicitacao)
        # sr.type_keys(serie_value)
        # sr.type_keys(serie)
        # data_emissao.type_keys(data_emissao_value)
        # vlr_nf.type_keys(valor_value)
        time.sleep(1)
        if inss is not None or irff is not None or piscofinscsl is not None or iss is not None:
            if inss > 0 or irff > 0 or piscofinscsl > 0 or iss > 0:
                detalhe_nota = janela.child_window(title="Valores da Nota", class_name="TTabSheet")
                irff_detail = detalhe_nota.child_window(class_name="TOvcDbPictureField", found_index=11)
                try:
                    self.wait_until_interactive(irff_detail)
                except TimeoutError as e:
                    print(str(e))
                    return
                irff_detail.click_input()
                time.sleep(1)
                pyautogui.press("tab")
                time.sleep(3)
                self.calculo_tributos(inss, irff, piscofinscsl, iss, valor_value)
        time.sleep(1)
        try:
            self.wait_until_interactive(observacao)
        except TimeoutError as e:
            print(str(e))
            return
        observacao.click_input()
        time.sleep(1)
        pyautogui.typewrite(obs)
        natureza_financeira_list = ['ALUGUEIS A PAGAR', 'COMBUSTÍVEIS/ LUBRIFICANTES', 'CUSTO COMBUSTIVEL NOVOS', 'CUSTO DESPACHANTE NOVOS', 'CUSTO FRETE NOVOS', 'CUSTO NOVOS', 'CUSTO OFICINA', 'DESP DESPACHANTE NOVOS', 'DESP FRETE VEIC NOVOS', 'DESP. COM SERVICOS DE OFICINA', 'DESPESA COM LAVACAO', 'DESPESA OFICINA', 'ENERGIA ELETRICA', 'FRETE', 'HONORARIO PESSOA JURIDICA', 'INFORMATICA HARDWARE', 'INFORMATICA SOFTWARE', 'INTERNET', 'MANUT. E CONSERV. DE', 'MATERIAL DE OFICINA DESPESA', 'PONTO ELETRONICO RA', 'SALARIO ERIBERTO', 'SALARIO MAGU', 'SALARIO VIVIANE', 'SALARIOS RA', 'SERVICO DE TERCEIROS FUNILARIA', 'SERVICOS DE TERCEIRO OFICINA', 'SOFTWARE', 'VALE TRANSPORTE RA', 'VD SALARIO', 'VD VEICULOS NOVOS', 'VIAGENS E ESTADIAS', 'AÇÕES EXTERNAS', 'AÇÕES LOJA', 'ADESIVOS', 'AGENCIA', 'BRINDES E CORTESIAS', 'DECORAÇÃO', 'DESENSOLVIMENTO SITE', 'DESP MKT CHERY FLORIPA', 'DISPARO SMS/WHATS', 'EVENTOS', 'EXPOSITORES', 'FACEBOOK', 'FACEBOOK/INSTAGRAM', 'FEE MENSAL', 'FEIRA/EVENTOS', 'FEIRAO', 'FOLLOWISE (MKT)', 'GOOGLE', 'INFLUENCIADORES', 'INSTITUCIONAL', 'INTEGRADOR (MKT)', 'JORNAL', 'LANCAMENTOS', 'LED', 'MARKETING', 'MERCADO LIVRE', 'MIDIA ON OUTROS', 'MIDIA/ONLINE', 'MKT', 'OUTDOOR', 'PANFLETOS', 'PATROCINIO', 'PORTAL GERACAO', 'PROSPECÇÃO', 'PUBLICIDADE E PROPAGANDA', 'RADIO', 'RD (MKT)', 'REGISTRO SITE', 'SISTEMAS (MKT)', 'SYONET AUTOMOVEIS', 'TELEVISAO', 'VENDAS EXTERNAS', 'VIDEOS', 'VITRINE', 'WISE (MKT)']
        if natureza_financeira_value in natureza_financeira_list:
            time.sleep(1)
            check_pis_cofins = janela.child_window(class_name="TCheckBox", found_index=0)
            try:
                self.wait_until_interactive(check_pis_cofins)
            except TimeoutError as e:
                print(str(e))
                return
            check_pis_cofins.click_input()
            # time.sleep(2)
            tab_natureza_credito = janela.child_window(title='Natureza Créditos Pis/Cofins', control_type='TabItem')
            try:
                self.wait_until_interactive(tab_natureza_credito)
            except TimeoutError as e:
                print(str(e))
                return
            tab_natureza_credito.click_input()
            time.sleep(1)
            nat_text = janela.child_window(class_name='TwwDBLookupCombo', found_index=1)
            try:
                self.wait_until_interactive(nat_text)
            except TimeoutError as e:
                print(str(e))
                return
            nat_text.click_input()
            time.sleep(1)
            pyautogui.typewrite('Aquisicao de bens utilizados como insumo')
            time.sleep(2)
            pyautogui.press('tab')
        # se for nota fiscal de produto 
        # habilitar_livro = janela.child_window(title='Quero esta nota no livro fiscal')
        # time.sleep(2)
        # habilitar_livro.click_input()
        # chave_acesso = self.get_chave_acesso(num_nf_value, id_solicitacao)
        # modelo = self.get_modelo(chave_acesso)
        # time.sleep(2)
        # janela_chave_modelo = janela.child_window(title="Modelo Fiscal / Chave e Outros")
        # janela_chave_modelo.click_input()
        # time.sleep(1)
        # modelo_fiscal = janela.child_window(class_name='TwwDBLookupCombo', found_index=0)
        # modelo_fiscal.click_input()
        # time.sleep(1)
        # if modelo == '55':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('nota fiscal eletronica')
        # elif modelo == '10':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('Conhecimento Aereo')
        # elif modelo == '67':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('Conhecimento de transp eletronico')
        # elif modelo == '9':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('Conhecimento de Transporte Aquaviario de cargas')
        # elif modelo == '57':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('Conhecimento de transporte eletronico - CT-e')
        # elif modelo == '11':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('Conhecimento de Transporte Ferroviario de Cargas')
        # elif modelo == '8':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('Conhecimento de Transporte Rodoviario de cargas')
        # elif modelo == '1B':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('NOTA FISCAL AVULSA')
        # elif modelo == '66':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('NOTA FISCAL DE ENERGIA ELETRICA ELETRONICA')
        # elif modelo == '29':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('Nota Fiscal/Conta de Fornecimento')
        # elif modelo == '1':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('NotaFiscal')
        # elif modelo == '21':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('NotaFiscal de Servico de Comunicacao')
        # elif modelo == '22':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('NotaFiscal de Servico de Telecomunicacoes')
        # elif modelo == '7':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('NotaFiscal de Servico de Transporte')
        # elif modelo == '6':
        #     modelo_fiscal.click_input()
        #     time.sleep(1)
        #     pyautogui.typewrite('NotaFiscal/Conta de Energia Eletrica')
        # time.sleep(1)
        # chave_nfe = janela.child_window(class_name='TDBEdit', found_index=6)
        # chave_nfe.type_keys(chave_acesso)
        time.sleep(1)
        try:
            self.wait_until_interactive(aba_cfop)
        except TimeoutError as e:
            print(str(e))
            return
        aba_cfop.click_input()
        time.sleep(2)
        CFOP_MAPPING = {
            '5102': '1556',
            '6102': '2556',
            '5403': '1407',
            '5656': '1407',
            '5405': '1407'
        }
        default_value = '2556' if estado != 'SC' else '1556'
        cfop_results = self.get_data_from_xml(chave_acesso)

        grouped_by_nature = {}
        for item in cfop_results:
            cfop = item['CFOP']
            valor_float = float(item['vProd'].replace(',', '.'))
            mapped_value = CFOP_MAPPING.get(cfop, default_value)

            if mapped_value in grouped_by_nature:
                grouped_by_nature[mapped_value] += valor_float
            else:
                grouped_by_nature[mapped_value] = valor_float

        valor_total_float = float(valor_value.replace(',', '.'))
        valor_already_added = 0

        for idx, (mapped_value, total) in enumerate(grouped_by_nature.items()):
            # Se não for o último item, use o valor total desse código de natureza
            if idx != len(grouped_by_nature) - 1:
                valor = round(total, 2)
                valor_already_added += valor
            else: # Se for o último item, ajuste para o total da nota
                valor = valor_total_float - valor_already_added
                
            valor_str = str(valor).replace('.', ',')
            self.execute_rpa_actions(mapped_value, valor_str)

        ##### se for nota fiscal de produto
        try:
            self.wait_until_interactive(tab_contab)
        except TimeoutError as e:
            print(str(e))
            return
        tab_contab.click_input()
        excluir = r"C:\Users\user\Documents\RPA_Project\imagens\Excluir.PNG"
        time.sleep(1)
        cod_contab = janela.child_window(class_name="TOvcNumericField", found_index=0)
        try:
            self.wait_until_interactive(cod_contab)
        except TimeoutError as e:
            print(str(e))
            return
        cod_contab.click_input()
        time.sleep(1)
        pyautogui.doubleClick()
        time.sleep(1)
        pyautogui.press("backspace")
        time.sleep(1)
        cod_contab.type_keys(codigo_contab)
        time.sleep(1)
        self.click_specific_button(raio)
        time.sleep(1)
        alterar = r"C:\Users\user\Documents\RPA_Project\imagens\Alterar.PNG"
        self.click_specific_button(alterar)
        time.sleep(2)
        self.get_conta_contabil()
        conta_contabil = self.get_conta_contabil()
        # time.sleep(2)
        self.fechar_conta_contabilizacao()
        informacao_dialog = janela.child_window(title="Informação")
        ok_button = informacao_dialog.child_window(class_name="Button")
        time.sleep(2)
        loop_continue = True
        while loop_continue:
            self.click_specific_button(excluir)
            time.sleep(1)
            try:
                self.wait_until_interactive(descricao)
            except TimeoutError as e:
                print(str(e))
                return
            descricao.click_input()
            for elem in janela.children():
                if "Informação" in elem.window_text():
                    ok_button.click()
                    loop_continue = False
                    break
        time.sleep(2)
        self.inserir_rateio(rateios, conta_contabil, valor_sg, valor_value, usa_rateio_centro_custo, rateios_aut)
        time.sleep(2)
        try:
            self.wait_until_interactive(tab_faturamento)
        except TimeoutError as e:
            print(str(e))
            return
        tab_faturamento.click_input()
        time.sleep(1)
        if tipo_pagamento_value == 'B':
            total_parcelas.type_keys(total_parcelas_value)
            if total_parcelas_value > 1:
                intervalo.type_keys("30")
        elif tipo_pagamento_value != 'B':
            num_parcela_n_boleto = len(num_parcelas)
            total_parcelas.type_keys(num_parcela_n_boleto)
            if num_parcela_n_boleto > 1:
                intervalo.type_keys("30")
        time.sleep(1)
        self.click_specific_button(gerar)
        time.sleep(1)
        try:
            self.wait_until_interactive(tipo_pagamento)
        except TimeoutError as e:
            print(str(e))
            return
        tipo_pagamento.click_input()
        if tipo_pagamento_value in ('B','A'):
            pyautogui.typewrite('Boleto Bancario')
            pyautogui.press('tab')
        elif tipo_pagamento_value in ('P','D'):
            pyautogui.typewrite('Deposito')
            pyautogui.press('tab')
        elif tipo_pagamento_value in ('E', 'C'):
            pyautogui.typewrite('Especie')
            pyautogui.press('tab')
        try:
            self.wait_until_interactive(natureza_despesa)
        except TimeoutError as e:
            print(str(e))
            return
        natureza_despesa.click_input()
        time.sleep(1)
        pyautogui.typewrite(natureza_financeira_value)
        time.sleep(1)
        pyautogui.press('tab')
        time.sleep(2)
        if tipo_pagamento_value == 'B':
            for boleto in boletos:
                vencimento.click_input()
                time.sleep(1)
                pyautogui.doubleClick()
                time.sleep(1)
                pyautogui.press("backspace")
                vencimento.type_keys(boleto[0])
                time.sleep(1)
                self.click_specific_button(barra_boleto)
                pyautogui.press("down")
        else:
            for n in num_parcelas:
                vencimento.click_input()
                time.sleep(1)
                pyautogui.doubleClick()
                time.sleep(1)
                pyautogui.press("backspace")
                vencimento.type_keys(n[0])
                time.sleep(1)
                self.click_specific_button(barra_boleto)
                pyautogui.press("down")
        print(terceiro)
        if terceiro == 'S':
            self.janela_se_terceiro(numeroos)
        time.sleep(2)
        submit_button.click_input()
        time.sleep(2)
        if estado != 'SC':
            pyautogui.press('tab')
            pyautogui.press('enter')
        time.sleep(1)
        start_time = time.time()
        while time.time() - start_time < 30:  
            for elem in janela.children():
                if "Aviso NF-e" in elem.window_text():
                    pyautogui.press("enter")
                    break
            time.sleep(0.5)  


    def execute_rpa_actions(self, mapped_value, valor):
        title = "Entrada Diversas / Operação: 52-Entrada Diversas"
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de nf")
            return 
        janela = app[title]
        inserir_natureza = r"C:\Users\user\Documents\RPA_Project\imagens\inserir_natureza.PNG"
        self.click_specific_button(inserir_natureza)
        time.sleep(2)
        pyautogui.typewrite(mapped_value)
        time.sleep(2)
        pesquisa_natureza = r"C:\Users\user\Documents\RPA_Project\imagens\pesquisa_natureza.PNG" 
        self.click_specific_button(pesquisa_natureza)
        time.sleep(1)
        for _ in range(2):
            pyautogui.press('tab')
        pyautogui.press('enter')
        time.sleep(2)
        outros = janela.child_window(class_name='TOvcDbPictureField', found_index=15)
        time.sleep(2)
        outros.click_input()
        for _ in range(2):
            pyautogui.press('tab')
            time.sleep(0.5)
        time.sleep(2)
        pyautogui.typewrite(valor)
        time.sleep(2)
        adicionar_cfop = r"C:\Users\user\Documents\RPA_Project\imagens\adicionar_cfop.PNG"
        self.click_specific_button(adicionar_cfop)

    def janela_se_terceiro(self, numeroos):
        title = "Entrada Diversas / Operação: 52-Entrada Diversas"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de cadastro_nf não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de nf")
            return 
        janela = app[title] 
        relacao_os = janela.child_window(control_type="TabItem", title='Relação de OS')
        cadeado = r"C:\Users\user\Documents\RPA_Project\imagens\Cadeado.PNG"
        pesquisa_os = r"C:\Users\user\Documents\RPA_Project\imagens\Procura_Os.PNG"
        numero_os = janela.child_window(class_name="TOvcNumericField", found_index=0)
        try:
            self.wait_until_interactive(numero_os)
        except TimeoutError as e:
            print(str(e))
            return
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
        title = "Entrada Diversas / Operação: 52-Nota de despesas Diversas"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de cadastro_nf não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de nota de despesas")
            return 
        janela = app[title]
        tab_dados = janela.child_window(control_type="TabItem", found_index=1)
        try:
            self.wait_until_interactive(tab_dados)
        except TimeoutError as e:
            print(str(e))
            return
        tab_dados.click_input()
        pyautogui.press("enter")
    
    def get_controle_value(self):
        title = "Entrada Diversas / Operação: 52-Nota de despesas Diversas"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de cadastro_nf não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de nota de despesas")
            return 
        janela = app[title]
        controle = janela.child_window(class_name="TDBEdit", found_index=2)
        properties = controle.legacy_properties()
        return properties['Value']
    
    def janela_imprimir_nota(self):
        title = "Ficha de Controle de Pagamento"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de ficha de controle de pagamento não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de imprimir nota")
            return 
        janela = app[title]
        # janela.wait('visible', timeout=120)
        imprimir = r"C:\Users\user\Documents\RPA_Project\imagens\Imprimir.PNG"
        self.click_specific_button(imprimir)

    def janela_secundario_imprimir_nota(self):
        title = "Report Destination"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de report destination não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela secundaria de imprimir nota")
            return 
        janela = app[title]
        # janela.wait('visible', timeout=120)
        # Declare Variables
        v_print = janela.child_window(found_index=12)
        try:
            self.wait_until_interactive(v_print)
        except TimeoutError as e:
            print(str(e))
            return
        # Set Comands
        v_print.click_input()

    def extract_pdf(self):
        title = "Ace Viewer"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de ace viewer não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Ace Viewer")
            return 
        janela = app[title]
        pyautogui.hotkey('alt', 'f')
        time.sleep(1)
        pyautogui.press('s')

    def save_as(self, num_docto, id_solicitacao):
        title = "Salvar como"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de salvar como não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Salvar como")
            return 
        janela = app[title]
        try:
            get_document = janela.child_window(title="APs (fixo)")
            get_document.click_input()
        except:
            pass
        time.sleep(2)
        file_name = janela.child_window(class_name="Edit")
        file_type = janela.child_window(class_name="AppControlHost", found_index=1)
        # Set Comands
        try:
            self.wait_until_interactive(file_name)
        except TimeoutError as e:
            print(str(e))
            return
        file_name.type_keys(f"AP_{num_docto}{id_solicitacao}")
        pyautogui.press('tab')
        try:
            self.wait_until_interactive(file_type)
        except TimeoutError as e:
            print(str(e))
            return
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

    def click_specific_button(self, button_image_path, confidence_level=0.8):
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
        title = "Ace Viewer"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de ace viewer não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Ace Viewer")
            return 
        janela = app[title]
        # janela.wait('visible', timeout=120)
        janela.set_focus()
        fechar = janela.child_window(title="Fechar")
        try:
            self.wait_until_interactive(fechar)
        except TimeoutError as e:
            print(str(e))
            return
        fechar.click_input()

    def close_aplications_half(self):
        janela_anterior_nbs_entrada = r"C:\Users\user\Documents\RPA_Project\imagens\Janela.PNG"
        time.sleep(2)
        self.click_specific_button(janela_anterior_nbs_entrada)

    def close_aplications_end(self):
        janela_anterior_sistema_financeiro = r"C:\Users\user\Documents\RPA_Project\imagens\Sair.PNG"
        self.click_specific_button(janela_anterior_sistema_financeiro)
    
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
        title = "Barra de Tarefas"
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Barra de Tarefas")
            return 
        janela = app[title]
        minimizar_google = janela.child_window(title="Google Chrome - 1 executando o windows")
        try:
            self.wait_until_interactive(minimizar_google)
        except TimeoutError as e:
            print(str(e))
            return
        minimizar_google.click_input()
        
    def fechar_conta_contabilizacao(self):
        title = "Alterar Conta de Contabilização"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de alterar conta contabilizacao não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de alterar conta de contabilizacao")
            return 
        janela = app[title]
        fechar = janela.child_window(class_name='TBitBtn', found_index=0)
        try:
            self.wait_until_interactive(fechar)
        except TimeoutError as e:
            print(str(e))
            return
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
        tamanho_aut = len(rateios_aut)
        tamanho_manual = len(rateios)
        if usa_rateio_centro_custo == 'N':
            for indice, rateio in enumerate(rateios):
                self.click_specific_button(incluir)
                time.sleep(1)
                contacontabil.type_keys(conta_contabil)
                pyautogui.press("tab")
                time.sleep(2)
                try:
                    centro_custo.type_keys(rateio[0]) 
                except:
                    pass
                time.sleep(1)
                historico_padrao.click_input()
                pyautogui.typewrite("85")
                vlr_rateio = self.get_valor_rateio_manual(rateio[1], valor_sg, vlr_nf)
                if indice != tamanho_manual - 1:
                    valor_rateio.type_keys(vlr_rateio)
                time.sleep(1)
                confirmar.click_input()
                time.sleep(1)
        elif usa_rateio_centro_custo == 'S':
            for indice, rateio_aut in enumerate(rateios_aut):
                self.click_specific_button(incluir)
                contacontabil.type_keys(conta_contabil)
                pyautogui.press("tab")
                time.sleep(2)
                try:
                    centro_custo.type_keys(rateio_aut[0]) 
                except:
                    pass
                time.sleep(1)
                historico_padrao.click_input()
                pyautogui.typewrite("85")
                vlr_rateio = self.get_valor_rateio_aut(vlr_nf, rateio_aut[1])
                if indice != tamanho_aut - 1:
                    valor_rateio.type_keys(vlr_rateio)
                time.sleep(1)
                confirmar.click_input()
                time.sleep(1)

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
        title = "Alterar Conta de Contabilização"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de alterar conta contabilizacao não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de alterar conta de contabilizacao")
            return 
        janela = app[title]
        time.sleep(2)
        conta_contabil = janela.child_window(class_name="TNBSContaContab")
        conta_contabil_value = conta_contabil.legacy_properties()
        print(conta_contabil_value['Value'])
        return conta_contabil_value['Value']
    
    
    def calculo_tributos(self, inss, irff, piscofinscsl, iss, nota_fiscal):
        title = ".:Cálculo de tributos:."
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de calculo de tributos não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de calculo de tributos")
            return 
        janela = app[title]
        check_csll = janela.child_window(class_name="TCheckBox", found_index=4)
        check_cofins = janela.child_window(class_name="TCheckBox", found_index=3)
        check_pis = janela.child_window(class_name="TCheckBox", found_index=2)
        confirm = janela.child_window(class_name="TBitBtn", found_index=1)
        iss_base_value = janela.child_window(class_name="TOvcNumericField", found_index=2)
        iss_value = janela.child_window(class_name="TOvcNumericField", found_index=3)
        irff_base_value = janela.child_window(class_name="TOvcNumericField", found_index=10)
        irff_value = janela.child_window(class_name="TOvcNumericField", found_index=11)
        inss_base_value = janela.child_window(class_name="TOvcNumericField", found_index=0)
        inss_value = janela.child_window(class_name="TOvcNumericField", found_index=1)
        if inss is not None:
            if inss > 0:
                inss_base_value.type_keys(nota_fiscal)
                inss_value.type_keys(str(inss).replace('.', ','))
        if irff is not None:
            if irff > 0:
                irff_base_value.type_keys(nota_fiscal)
                irff_value.type_keys(str(irff).replace('.', ','))
        if piscofinscsl is not None:
            if piscofinscsl > 0:
                check_csll.click_input()
                time.sleep(0.5)
                check_cofins.click_input()
                time.sleep(0.5)
                check_pis.click_input()
                time.sleep(0.5)
        if iss is not None:
            if iss > 0:
                iss_base_value.type_keys(nota_fiscal)
                iss_value.type_keys(str(iss).replace('.', ','))
        time.sleep(1)
        confirm.click_input()
    
    def extract_text_from_pdf(self, num_docto, id_solicitacao):
        pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
        imagens = convert_from_path(caminho_pdf)
        texto_total = ""
        for imagem in imagens:
            texto = pytesseract.image_to_string(imagem)
            if texto:
                texto_total += texto + '\n'
        return texto_total

    def find_isolated_5102(self, text):
        pattern = r'(?<!\d)5102(?!\d)'
        result = re.search(pattern, text)
        return True if result else False
    
    def get_chave_acesso(self, num_docto, id_solicitacao):
        # pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        # caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
        # imagens = convert_from_path(caminho_pdf)
        # texto_total = ""
        # for imagem in imagens:
        #     texto = pytesseract.image_to_string(imagem)
        #     if texto:
        #         texto_total += texto + '\n'
        # chave_acesso_match = re.search(r'((\d{4}[\W\s]){10}\d{4})', texto_total)
        # chave_acesso = chave_acesso_match.group(0) if chave_acesso_match else None
        # if chave_acesso:
        #     nao_numericos = re.findall(r'[^\d]', chave_acesso)
        #     if nao_numericos:
        #         chave_acesso = re.sub(r'[^\d]', '', chave_acesso)
        # return chave_acesso
        pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
        imagens = convert_from_path(caminho_pdf)
        texto_total = ""
        for imagem in imagens:
            texto = pytesseract.image_to_string(imagem)
            if texto:
                texto_total += texto + '\n'
        # Primeiro procuramos pelo novo padrão (44 dígitos consecutivos)
        chave_acesso_match = re.search(r'(\d{44})', texto_total)
        # Se não encontrarmos, procuramos pelo padrão anterior
        if not chave_acesso_match:
            chave_acesso_match = re.search(r'((\d{4}[\W\s]){10}\d{4})', texto_total)
        chave_acesso = chave_acesso_match.group(0) if chave_acesso_match else None
        if chave_acesso:
            nao_numericos = re.findall(r'[^\d]', chave_acesso)
            if nao_numericos:
                chave_acesso = re.sub(r'[^\d]', '', chave_acesso)
        return chave_acesso

    def get_modelo(self, chave_acesso):
        return chave_acesso[20:22] if len(chave_acesso) > 21 else "Não foi possível extrair os números"
    
    def get_serie(self, num_docto, id_solicitacao):
        pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
        caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
        imagens = convert_from_path(caminho_pdf)
        texto_total = ""
        for imagem in imagens:
            texto = pytesseract.image_to_string(imagem)
            if texto:
                texto_total += texto + '\n'
        serie_match = re.search(r's[ée]rie\s*[:\.\-\s]?\s*(\d+)', texto_total, re.IGNORECASE)
        serie = serie_match.group(1) if serie_match else "Série não encontrada"
        return serie
    
    def get_valor_icms(self, num_docto, id_solicitacao):
        caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
        texto_total = ""    
        with pdfplumber.open(caminho_pdf) as pdf:
            for page in pdf.pages:
                texto = page.extract_text()
                if texto:
                    texto_total += texto + '\n'
        patterns = [
            r'(V\.|VALOR) TOTAL DOS PRODUTOS\S*\s+([\d\.,]+)\s+([\d\.,]+)',
            r'(V\.|VALOR) TOTAL PRODUTOS\S*\s+([\d\.,]+)\s+([\d\.,]+)'      
        ]
        for pattern in patterns:
            valor_icms_match = re.search(pattern, texto_total)
            if valor_icms_match:
                valor_string = valor_icms_match.group(3).replace('.', '').replace(',', '.')
                valor_icms = float(valor_string)
                return valor_icms
        else:
            return False
    
    def convert_to_date(self, number):
        str_num = str(number)
        day = int(str_num[:2])
        month = int(str_num[2:4])
        year = int(str_num[4:])
        return datetime.date(year, month, day) 
        
    def check_conditions(self, tipo_docto, chave_acesso, serie, natureza, vencimento, inss, irff, piscofinscsl, tipo_pagamento, icms, boletos):
        current_date = datetime.date.today()  # Pegando apenas a data, sem hora, minuto, segundo, etc.
        converted_vencimento = self.convert_to_date(vencimento)
        # if tipo_docto != "NFE":
        #     return False, "Tipo de documento inválido"
        if not chave_acesso:
            return False, "chave de acesso não encontrada"
        # if not serie:
        #     return False, "serie não encontrada"
        # if not icms or icms == 0.0:
        #     return False, "icms nao encontrado ou zerado"
        # if not natureza:
        #     return False, "natureza 5102 não encontrada"
        if tipo_pagamento == 'B':
            for boleto in boletos:
                converted_boleto_date = self.convert_to_date(boleto[0])
                if converted_boleto_date < current_date:
                    return False, "O primeiro boleto tem vencimento anterior à data atual"
                break 
        elif tipo_pagamento != 'B':
            if converted_vencimento < current_date:
                return False, "vencimento menor que data de efetivacao"
        if inss:
            return False, "inss encontrado"
        if irff:
            return False, "irff encontrado"
        if piscofinscsl:
            return False, "piscofinscsl encontrado"
        if tipo_pagamento not in ('B', 'A', 'P', 'D', 'E', 'C'):
            return False, "tipo de pagamento diferente do configurado"
        
        return True, "Condições aceitas"
        
    def send_message_pre_verification(self, msg, id_solicitacao, numerodocto):
        # chat_id = '-995541913' 
        chat_id_result = database.consultar_chat_id()
        token_result = database.consultar_token_bot()
        chat_id = chat_id_result[0][0]
        token = token_result[0][0]
        # base_url = f"https://api.telegram.org/bot6563290849:AAEVdZmaKFWOTmUdIYeAcGhq3_cUu4sX2qE/sendMessage"
        base_url = f"https://api.telegram.org/bot{token}/sendMessage"
        payload = {
            'chat_id': chat_id,
            'text': f"[ERROR]: {msg}, id_solicitacao: {id_solicitacao}, numero_docto: {numerodocto}"
        }
        response = requests.post(base_url, data=payload)
        if response.status_code != 200:
            print(f"Failed to send message. Response: {response.content}")
        else:
            print("Message sent successfully to Telegram!")

    def send_message_with_traceback(self, id_solicitacao, numerodocto):
        chat_id_result = database.consultar_chat_id()
        token_result = database.consultar_token_bot()
        chat_id = chat_id_result[0][0]
        token = token_result[0][0]
        error_message = traceback.format_exc()
        detailed_msg = f"[TRACEBACK ERROR]: {error_message}, id_solicitacao: {id_solicitacao}, numero_docto: {numerodocto}"
        # chat_id = '-995541913' 
        # base_url = f"https://api.telegram.org/bot6563290849:AAEVdZmaKFWOTmUdIYeAcGhq3_cUu4sX2qE/sendMessage"
        base_url = f"https://api.telegram.org/bot{token}/sendMessage"
        payload = {
            'chat_id': chat_id,
            'text': detailed_msg
        }
        response = requests.post(base_url, data=payload)
        if response.status_code != 200:
            print(f"Failed to send traceback message. Response: {response.content}")
        else:
            print("Traceback message sent successfully to Telegram!")

    def get_data_from_xml(self, chave_acesso):
        path = rf'C:\Users\user\Downloads\{chave_acesso}.xml'
        tree = ET.parse(path)
        root = tree.getroot()
        # Identifica o namespace
        ns = None
        for elem in root.iter():
            if '}' in elem.tag:
                ns = {'ns': elem.tag.split('}')[0].strip('{')}
                break
        if ns is None:
            raise ValueError("Namespace não encontrado no XML.")
        # Processa os elementos <det> para obter CFOP, vProd e vDesc
        cfop_values = {}
        det_elements = root.findall('.//ns:det', namespaces=ns)
        for det in det_elements:
            cfop_elem = det.find('.//ns:CFOP', namespaces=ns)
            vprod_elem = det.find('.//ns:vProd', namespaces=ns)
            vdesc_elem = det.find('.//ns:vDesc', namespaces=ns)
            if cfop_elem is not None and vprod_elem is not None:
                cfop = cfop_elem.text
                vprod = float(vprod_elem.text.replace(',', '.')) # Converte para float considerando a vírgula
                if vdesc_elem is not None:
                    vprod -= float(vdesc_elem.text.replace(',', '.')) # Desconta o valor vDesc se ele existir
                # Arredonda o valor para 2 casas decimais
                vprod = math.ceil(vprod * 100) / 100
                if cfop in cfop_values:
                    cfop_values[cfop] += vprod  # Soma o valor vProd ao CFOP existente
                else:
                    cfop_values[cfop] = vprod  # Cria uma nova entrada para o CFOP com o valor vProd
        # Converte o dicionário processado em uma lista de dicionários, formatando o valor para string
        results = [{'CFOP': key, 'vProd': str(value).replace('.', ',')} for key, value in cfop_values.items()]
        return results
    
    def funcao_main(self):
        registros = database.consultar_dados_cadastro()
        empresa_anterior = None
        for row in registros:
            # try:
            id_empresa = row[18]
            id_solicitacao = row[0]
            tipo_pagamento_value = row[5]
            notas_fiscais = database.consultar_nota_fiscal(id_solicitacao)
            if notas_fiscais:
                numerodocto = notas_fiscais[0][2]
                serie_value = notas_fiscais[0][3]
                data_emissao_value = notas_fiscais[0][1] 
                tipo_docto_value = notas_fiscais[0][5]
                valor_value = notas_fiscais[0][0]
                inss = notas_fiscais[0][6]
                irff = notas_fiscais[0][7]
                piscofinscsl = notas_fiscais[0][8]
                iss = notas_fiscais[0][9]
                vencimento_value = notas_fiscais[0][4]
            wise_instance = Wise()
            wise_instance.get_nf_values(numerodocto, id_solicitacao)
            time.sleep(2)
            wise_instance.save_as(numerodocto, id_solicitacao)
            time.sleep(1)
            chave_de_acesso_value = self.get_chave_acesso(numerodocto, id_solicitacao)
            print(chave_de_acesso_value)
            serie_nota = self.get_serie(numerodocto, id_solicitacao)
            pdf_inteiro = self.extract_text_from_pdf(numerodocto, id_solicitacao)
            natureza_value = self.find_isolated_5102(pdf_inteiro)
            icms = self.get_valor_icms(numerodocto, id_solicitacao)
            boletos = database.consultar_boleto(id_solicitacao)
            time.sleep(1)
            success, message = self.check_conditions(tipo_docto_value, chave_de_acesso_value, serie_nota, natureza_value, vencimento_value, inss, irff, piscofinscsl, tipo_pagamento_value, icms, boletos)  
            time.sleep(3)
            self.back_to_nbs()
            time.sleep(2)
            wise_instance.fechar_aba()
            time.sleep(5)
            time.sleep(1)
            # except: 
            #     self.send_message_with_traceback(id_solicitacao, numerodocto)
            if success:
                try:
                    wise_instance.get_xml(chave_de_acesso_value)
                    time.sleep(2)
                    self.back_to_nbs()
                    # wise_instance.fechar_aba()
                    # if success:
                        # try:
                    empresa_atual = row[2]
                    if empresa_atual != empresa_anterior:
                        self.close_aplications_end()
                        time.sleep(3)
                        self.open_application()
                        self.login()
                        self.janela_empresa_filial(row[2], row[3])
                    empresa_anterior = empresa_atual
                    print(empresa_anterior, empresa_atual)
                    self.access_contas_a_pagar()
                    self.janela_entrada()
                    self.importar_xml()
                    self.abrir_xml(chave_de_acesso_value)
                    cnpj = row[1]
                    contab_descricao_value = row[6]
                    cod_contab_value = row[7]
                    total_parcelas_value = row[9]
                    natureza_financeira_value = row[8]
                    usa_rateio_centro_custo = row[14] 
                    valor_sg = row[15] 
                    id_rateiocc = row[16]
                    obs = row[17]
                    rateios_aut = database.consultar_rateio_aut(id_rateiocc)
                    boletos = database.consultar_boleto(id_solicitacao)
                    rateios = database.consultar_rateio(id_solicitacao)
                    num_parcelas = database.numero_parcelas(id_solicitacao, numerodocto)
                    numeroos = row[11]
                    terceiro = row[12]
                    estado = row[13]
                    time.sleep(3)
                    self.janela_cadastro_nf(cnpj, numerodocto, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value, numeroos, terceiro, estado, usa_rateio_centro_custo, valor_sg, rateios, rateios_aut, inss, irff, piscofinscsl, iss, vencimento_value, obs, cod_contab_value, boletos, num_parcelas, id_solicitacao, chave_de_acesso_value)
                    self.janela_imprimir_nota()
                    self.janela_secundario_imprimir_nota()
                    self.extract_pdf()
                    self.save_as(numerodocto, id_solicitacao)
                    self.close_extract_pdf_window()
                    time.sleep(2)
                    self.click_on_cancel()
                    self.janela_valores()
                    time.sleep(2)
                    num_controle = self.get_controle_value()
                    time.sleep(2)
                    print(num_controle)
                    time.sleep(5)
                    wise_instance.Anexar_AP(id_solicitacao, num_controle, numerodocto)
                    time.sleep(2)
                    wise_instance.get_pdf_file(numerodocto, id_solicitacao)
                    time.sleep(4)
                    wise_instance.confirm()
                    time.sleep(3)
                    database.atualizar_anexosolicitacaogasto(numerodocto)
                    time.sleep(2)
                    send_mail.enviar_email(id_solicitacao, numerodocto)
                    time.sleep(4)
                    self.back_to_nbs()
                    time.sleep(2)
                    self.close_aplications_half()
                    time.sleep(3)
                except:
                    self.send_message_with_traceback(id_solicitacao, numerodocto)
            else:
                self.send_message_pre_verification(message, id_solicitacao, numerodocto)
                database.autoriza_rpa_para_n(id_solicitacao)
            
rpa = NbsRpa()
rpa.funcao_main()
