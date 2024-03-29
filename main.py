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
from decimal import Decimal, ROUND_HALF_UP
import xml.etree.ElementTree as ET
import math
import send_mail
from decimal import Decimal, ROUND_DOWN
import os
import nlp3
import pyperclip
from pathlib import Path

class NbsRpa():
    
    def __init__(self):
        self.existe_xml = False
        self.efetivado = False
    
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
        user.click_input()
        time.sleep(1)
        user.type_keys("NBS")
        time.sleep(2)
        try:
            self.wait_until_interactive(password)
        except TimeoutError as e:
            print(str(e))
            return
        password.click_input()
        time.sleep(1)
        password.type_keys("gerati2023")
        time.sleep(2)
        try:
            self.wait_until_interactive(server)
        except TimeoutError as e:
            print(str(e))
            return
        server.click_input()
        time.sleep(1)
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

    def janela_entrada(self, tipo_docto_value):
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
        time.sleep(2)
        if tipo_docto_value != 'NFE':
            button_image_path = r"C:\Users\user\Documents\RPA_Project\imagens\Inserir_Nota.PNG"
            time.sleep(1)
            self.click_specific_button(button_image_path)
        elif tipo_docto_value == 'NFE':
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

    def abrir_xml(self, chave_acesso, num_docto, id_solicitacao):
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
        time.sleep(5)
        if self.xml_existe == False:
            # file_name.type_keys(chave_acesso)
            self.type_slowly(file_name, chave_acesso)
        else:
            # file_name.type_keys(f"xml_{num_docto}_{id_solicitacao}")
            self.type_slowly(file_name, f"xml_{num_docto}_{id_solicitacao}")
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(3)
        pyautogui.typewrite('Hyundai_Diversos')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(2)
        aceitar = r"C:\Users\user\Documents\RPA_Project\imagens\aceitar.PNG"
        self.click_specific_button(aceitar)

    def janela_cadastro_nf(self, cpf_cnpj_value, num_nf_value, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value, numeroos, terceiro, estado, usa_rateio_centro_custo, valor_sg, rateios, rateios_aut, inss, irff, piscofinscsl, iss, vencimento_value, obs, codigo_contab, boletos, num_parcelas, id_solicitacao, chave_acesso, cod_matriz, cidade_cliente, empresa):
        title = "Entrada Diversas / Operação: 52-Entrada Diversas"
        if not self.esperar_janela_visivel(title, timeout=60):
            erro_msg = "Falha: Janela de cadastro_nf não está visível."
            print(erro_msg)
            raise Exception(erro_msg)
            # return
        time.sleep(5)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de nf")
            # return 
        janela = app[title]
        time.sleep(2)
        janela.set_focus()
        print(vencimento_value)
        cpf_cnpj = janela.child_window(class_name="TCPF_CGC")
        esp = janela.child_window(class_name="TOvcDbPictureField", found_index=2)
        livro = janela.child_window(title="Quero esta nota no livro fiscal")
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
        nota_ext = janela.child_window(title="Nota Extemporânea") 
        regime_norma = janela.child_window(title='Docto Fiscal emitido com base em Regime Esp. ou Norma Específica')
        vlr_iss = janela.child_window(class_name="TOvcDbPictureField", found_index=15)
        fisco_municipal = janela.child_window(title='Fisco Municipal')
        aliquota_field = janela.child_window(class_name='TOvcDbPictureField', found_index=14)
        # Set Comands
        if tipo_docto_value == 'NFS':
            try:
                self.wait_until_interactive(cpf_cnpj)
            except TimeoutError as e:
                print(str(e))
                return
            cpf_cnpj.type_keys(cpf_cnpj_value)
            cpf_cnpj.type_keys("{ENTER}")
            try:
                self.wait_until_interactive(esp)
            except TimeoutError as e:
                print(str(e))
                return
            esp.click_input()
            esp.type_keys(tipo_docto_value)
            try:
                self.wait_until_interactive(num_nf)
            except TimeoutError as e:
                print(str(e))
                return
            num_nf.click_input()
            time.sleep(1)
            num_nf.type_keys(num_nf_value)
            try:
                self.wait_until_interactive(sr)
            except TimeoutError as e:
                print(str(e))
                return
            # serie = self.get_serie(num_nf_value, id_solicitacao)
            sr.type_keys(serie_value)
            # sr.type_keys(serie)
            try:
                self.wait_until_interactive(data_emissao)
            except TimeoutError as e:
                print(str(e))
                return
            data_emissao.type_keys(data_emissao_value)
            try:
                self.wait_until_interactive(vlr_iss)
            except TimeoutError as e:
                print(str(e))
                return
            if serie_value not in ('TL', 'CO', 'CCF'):
                vlr_iss.type_keys(valor_value)
                if self.xml_existe == False:
                    aliquota = self.get_aliquota(num_nf_value, id_solicitacao)
                    aliquota_img = self.extract_aliquota_from_image(num_nf_value, id_solicitacao)
                time.sleep(3)
                try:
                    self.wait_until_interactive(aliquota_field)
                except TimeoutError as e:
                    print(str(e))
                    return
                time.sleep(1)
                aliquota_field.click_input()
                time.sleep(1)
                if self.xml_existe == False and self.get_aliquota_from_xml_nfs(num_nf_value, id_solicitacao) == None:
                    if aliquota != '0,00' and aliquota is not None:
                        time.sleep(1)
                        pyautogui.doubleClick()
                        time.sleep(1)
                        pyautogui.press('backspace')
                        time.sleep(1)
                        pyautogui.typewrite(aliquota)
                    elif aliquota is None and aliquota_img is not None:
                        time.sleep(1)
                        pyautogui.doubleClick()
                        time.sleep(1)
                        pyautogui.press('backspace')
                        time.sleep(1)
                        pyautogui.typewrite(aliquota_img)
                        time.sleep(1)
                elif (self.xml_existe == True or self.xml_existe == False) and self.get_aliquota_from_xml_nfs(num_nf_value, id_solicitacao) != None:
                    aliquota_xml = self.get_aliquota_from_xml_nfs(num_nf_value, id_solicitacao)
                    if aliquota_xml != '0,00' and aliquota_xml is not None:
                        time.sleep(1)
                        pyautogui.doubleClick()
                        time.sleep(1)
                        pyautogui.press('backspace')
                        time.sleep(1)
                        pyautogui.typewrite(aliquota_xml)
                for _ in range(3):
                    pyautogui.press('tab')
                    time.sleep(2)
                if serie_value != 'FA':
                    if self.get_item_lista_servico_from_xml_nfs(num_nf_value, id_solicitacao) is None:
                        cod_servico = self.get_service_code(num_nf_value, id_solicitacao)
                        cod_servico_img = self.extract_service_code_from_image(num_nf_value, id_solicitacao)
                        time.sleep(2)
                        if cod_servico is not None:
                            pyautogui.typewrite(cod_servico)
                        elif cod_servico is None and cod_servico_img is not None:
                            pyautogui.typewrite(cod_servico_img)
                        else:
                            cod_servico_nlp = nlp3.refine_best_match_code(num_nf_value, id_solicitacao)
                            pyautogui.typewrite(cod_servico_nlp)
                    elif self.get_item_lista_servico_from_xml_nfs(num_nf_value, id_solicitacao) != None:
                        cod_servico_xml = self.get_item_lista_servico_from_xml_nfs(num_nf_value, id_solicitacao)
                        if cod_servico_xml:
                            pyautogui.typewrite(cod_servico_xml)
                        else:
                            cod_servico_nlp = nlp3.refine_best_match_code(num_nf_value, id_solicitacao)
                            pyautogui.typewrite(cod_servico_nlp)
                else:
                    pyautogui.typewrite("1")
            try:
                self.wait_until_interactive(vlr_nf)
            except TimeoutError as e:
                print(str(e))
                return
            vlr_nf.type_keys(valor_value)
            time.sleep(1)
            if serie_value not in ('TL', 'CO', 'CCF'):
                time.sleep(2)
                fisco_municipal.click_input()
                for _ in range(3):
                    pyautogui.press('tab')
                    time.sleep(0.5)
                time.sleep(1)
            elif serie_value in ('TL', 'CO', 'CCF'):
                time.sleep(2)
                for _ in range(5):
                    pyautogui.press('tab')
                    time.sleep(0.5)
                time.sleep(1)
            texto_completo = f"{obs} Numero de solicitacao: {id_solicitacao}"
            pyautogui.typewrite(texto_completo)
            time.sleep(1)
            try:
                self.wait_until_interactive(nota_ext)
            except TimeoutError as e:
                print(str(e))
                return
            mes_data_emissao = int(data_emissao_value[2:4])
            mes_atual = datetime.datetime.now().month
            if mes_data_emissao != mes_atual:
                nota_ext.click_input()
            time.sleep(2) 

        elif tipo_docto_value not in ('NFS', 'NFE') and (serie_value in ('TL', 'CO', 'CCF') or serie_value is None):
            try:
                self.wait_until_interactive(cpf_cnpj)
            except TimeoutError as e:
                print(str(e))
                return
            cpf_cnpj.type_keys(cpf_cnpj_value)
            cpf_cnpj.type_keys("{ENTER}")
            try:
                self.wait_until_interactive(livro)
            except TimeoutError as e:
                print(str(e))
                return
            # esp.click_input()
            time.sleep(2)
            # esp.type_keys(tipo_docto_value)
            livro.click_input()
            try:
                self.wait_until_interactive(num_nf)
            except TimeoutError as e:
                print(str(e))
                return
            num_nf.click_input()
            time.sleep(1)
            num_nf.type_keys(num_nf_value)
            try:
                self.wait_until_interactive(sr)
            except TimeoutError as e:
                print(str(e))
                return
            # serie = self.get_serie(num_nf_value, id_solicitacao)
            try:
                sr.type_keys(serie_value)
            except:
                pass
            # sr.type_keys(serie)
            try:
                self.wait_until_interactive(data_emissao)
            except TimeoutError as e:
                print(str(e))
                return
            data_emissao.type_keys(data_emissao_value)
            try:
                self.wait_until_interactive(vlr_iss)
            except TimeoutError as e:
                print(str(e))
                return
            try:
                self.wait_until_interactive(vlr_nf)
            except TimeoutError as e:
                print(str(e))
                return
            vlr_nf.type_keys(valor_value)
            time.sleep(2)
            for _ in range(5):
                pyautogui.press('tab')
                time.sleep(0.5)
            time.sleep(1)
            texto_completo = f"{obs} Numero de solicitacao: {id_solicitacao}"
            pyautogui.typewrite(texto_completo)
            time.sleep(1)
            try:
                self.wait_until_interactive(nota_ext)
            except TimeoutError as e:
                print(str(e))
                return
            mes_data_emissao = int(data_emissao_value[2:4])
            mes_atual = datetime.datetime.now().month
            if mes_data_emissao != mes_atual:
                nota_ext.click_input()
            time.sleep(2) 

        elif tipo_docto_value in ('NFS', 'NFE') and serie_value is None:
            try:
                self.wait_until_interactive(cpf_cnpj)
            except TimeoutError as e:
                print(str(e))
                return
            cpf_cnpj.type_keys(cpf_cnpj_value)
            cpf_cnpj.type_keys("{ENTER}")
            try:
                self.wait_until_interactive(esp)
            except TimeoutError as e:
                print(str(e))
                return
            esp.click_input()
            esp.type_keys(tipo_docto_value)
            try:
                self.wait_until_interactive(num_nf)
            except TimeoutError as e:
                print(str(e))
                return
            num_nf.click_input()
            time.sleep(1)
            num_nf.type_keys(num_nf_value)
            try:
                self.wait_until_interactive(sr)
            except TimeoutError as e:
                print(str(e))
                return
            # serie = self.get_serie(num_nf_value, id_solicitacao)
            try:
                sr.type_keys(serie_value)
            except:
                pass
            # sr.type_keys(serie)
            try:
                self.wait_until_interactive(data_emissao)
            except TimeoutError as e:
                print(str(e))
                return
            data_emissao.type_keys(data_emissao_value)
            try:
                self.wait_until_interactive(vlr_iss)
            except TimeoutError as e:
                print(str(e))
                return
            try:
                self.wait_until_interactive(vlr_nf)
            except TimeoutError as e:
                print(str(e))
                return
            vlr_nf.type_keys(valor_value)
            time.sleep(2)
            for _ in range(5):
                pyautogui.press('tab')
                time.sleep(0.5)
            time.sleep(1)
            texto_completo = f"{obs} Numero de solicitacao: {id_solicitacao}"
            pyautogui.typewrite(texto_completo)
            time.sleep(1)
            try:
                self.wait_until_interactive(nota_ext)
            except TimeoutError as e:
                print(str(e))
                return
            mes_data_emissao = int(data_emissao_value[2:4])
            mes_atual = datetime.datetime.now().month
            if mes_data_emissao != mes_atual:
                nota_ext.click_input()
            time.sleep(2) 
        elif tipo_docto_value not in ('NFS', 'NFE') and serie_value in ('FA'):
            try:
                self.wait_until_interactive(cpf_cnpj)
            except TimeoutError as e:
                print(str(e))
                return
            cpf_cnpj.type_keys(cpf_cnpj_value)
            cpf_cnpj.type_keys("{ENTER}")
            try:
                self.wait_until_interactive(livro)
            except TimeoutError as e:
                print(str(e))
                return
            # esp.click_input()
            time.sleep(2)
            # esp.type_keys(tipo_docto_value)
            livro.click_input()
            try:
                self.wait_until_interactive(num_nf)
            except TimeoutError as e:
                print(str(e))
                return
            num_nf.click_input()
            time.sleep(1)
            num_nf.type_keys(num_nf_value)
            try:
                self.wait_until_interactive(sr)
            except TimeoutError as e:
                print(str(e))
                return
            # serie = self.get_serie(num_nf_value, id_solicitacao)
            sr.type_keys(serie_value)
            # sr.type_keys(serie)
            try:
                self.wait_until_interactive(data_emissao)
            except TimeoutError as e:
                print(str(e))
                return
            data_emissao.type_keys(data_emissao_value)
            try:
                self.wait_until_interactive(vlr_iss)
            except TimeoutError as e:
                print(str(e))
                return
            vlr_iss.type_keys(valor_value)
            if self.xml_existe == False:
                aliquota = self.get_aliquota(num_nf_value, id_solicitacao)
                aliquota_img = self.extract_aliquota_from_image(num_nf_value, id_solicitacao)
            time.sleep(3)
            try:
                self.wait_until_interactive(aliquota_field)
            except TimeoutError as e:
                print(str(e))
                return
            time.sleep(1)
            aliquota_field.click_input()
            time.sleep(1)
            if self.xml_existe == False and self.get_aliquota_from_xml_nfs(num_nf_value, id_solicitacao) == None:
                if aliquota != '0,00' and aliquota is not None:
                    time.sleep(1)
                    pyautogui.doubleClick()
                    time.sleep(1)
                    pyautogui.press('backspace')
                    time.sleep(1)
                    pyautogui.typewrite(aliquota)
                elif aliquota is None and aliquota_img is not None:
                    time.sleep(1)
                    pyautogui.doubleClick()
                    time.sleep(1)
                    pyautogui.press('backspace')
                    time.sleep(1)
                    pyautogui.typewrite(aliquota_img)
                    time.sleep(1)
            elif (self.xml_existe == True or self.xml_existe == False) and self.get_aliquota_from_xml_nfs(num_nf_value, id_solicitacao) != None:
                aliquota_xml = self.get_aliquota_from_xml_nfs(num_nf_value, id_solicitacao)
                if aliquota_xml != '0,00' and aliquota_xml is not None:
                    time.sleep(1)
                    pyautogui.doubleClick()
                    time.sleep(1)
                    pyautogui.press('backspace')
                    time.sleep(1)
                    pyautogui.typewrite(aliquota_xml)
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(3)
                pyautogui.typewrite('1')   
            try:
                self.wait_until_interactive(vlr_nf)
            except TimeoutError as e:
                print(str(e))
                return
            vlr_nf.type_keys(valor_value)
            time.sleep(2)
            fisco_municipal.click_input()
            for _ in range(3):
                pyautogui.press('tab')
                time.sleep(0.5)
            time.sleep(1)
            texto_completo = f"{obs} Numero de solicitacao: {id_solicitacao}"
            pyautogui.typewrite(texto_completo)
            time.sleep(1)
            try:
                self.wait_until_interactive(nota_ext)
            except TimeoutError as e:
                print(str(e))
                return
            mes_data_emissao = int(data_emissao_value[2:4])
            mes_atual = datetime.datetime.now().month
            if mes_data_emissao != mes_atual:
                nota_ext.click_input()
            time.sleep(2) 

        if tipo_docto_value == 'NFE':
            try:
                self.wait_until_interactive(observacao)
            except TimeoutError as e:
                print(str(e))
                return
            observacao.click_input()
            time.sleep(1)
            texto_completo = f"{obs} Numero de solicitacao: {id_solicitacao}"
            pyautogui.typewrite(texto_completo)
            time.sleep(1)
            try:
                self.wait_until_interactive(nota_ext)
            except TimeoutError as e:
                print(str(e))
                return
            if tipo_docto_value == 'NFE':
                data_emissao_xml_value = self.get_data_emissao_from_xml(chave_acesso, num_nf_value, id_solicitacao)
                mes_xml = int(data_emissao_xml_value.split('-')[1])
                mes_atual = datetime.datetime.now().month
                if mes_xml != mes_atual:
                    nota_ext.click_input()
                time.sleep(2)

            cfop_results = self.get_data_from_xml(chave_acesso, num_nf_value, id_solicitacao)
            time.sleep(1)
            numero_serie = self.get_serie_from_xml(chave_acesso, num_nf_value, id_solicitacao)
            has_cfop_starting_with_59 = any(item['CFOP'].startswith("59") for item in cfop_results)

            if numero_serie in ('890', '891','892','893','894','895','896','897','898','899') or has_cfop_starting_with_59:
                regime_norma.click_input()
            time.sleep(2)
        if serie_value not in ('TL','CO', 'CCF'):
            natureza_financeira_list = ['OUTROS MATERIAIS GRAFICOS', 'ALUGUEIS A PAGAR', 'COMBUSTIVEIS/ LUBRIFICANTES', 'CUSTO COMBUSTIVEL NOVOS', 'CUSTO DESPACHANTE NOVOS', 'CUSTO FRETE NOVOS', 'CUSTO NOVOS', 'CUSTO OFICINA', 'DESP DESPACHANTE NOVOS', 'DESP FRETE VEIC NOVOS', 'DESP. COM SERVICOS DE OFICINA', 'DESPESA COM LAVACAO', 'DESPESA OFICINA', 'ENERGIA ELETRICA', 'FRETE', 'HONORARIO PESSOA JURIDICA', 'INFORMATICA HARDWARE', 'INFORMATICA SOFTWARE', 'INTERNET', 'MANUT. E CONSERV. DE', 'MATERIAL DE OFICINA DESPESA', 'PONTO ELETRONICO RA', 'SALARIO ERIBERTO', 'SALARIO MAGU', 'SALARIO VIVIANE', 'SALARIOS RA', 'SERVICO DE TERCEIROS FUNILARIA', 'SERVICOS DE TERCEIRO OFICINA', 'SOFTWARE', 'VALE TRANSPORTE RA', 'VD SALARIO', 'VD VEICULOS NOVOS', 'VIAGENS E ESTADIAS', 'AÇÕES EXTERNAS', 'AÇÕES LOJA', 'ADESIVOS', 'AGENCIA', 'BRINDES E CORTESIAS', 'DECORAÇÃO', 'DESENSOLVIMENTO SITE', 'DESP MKT CHERY FLORIPA', 'DISPARO SMS/WHATS', 'EVENTOS', 'EXPOSITORES', 'FACEBOOK', 'FACEBOOK/INSTAGRAM', 'FEE MENSAL', 'FEIRA/EVENTOS', 'OUTROS EVENTOS', 'OUTROS (MKT)', 'FEIRAO', 'FOLLOWISE (MKT)', 'GOOGLE', 'INFLUENCIADORES', 'INSTITUCIONAL', 'INTEGRADOR (MKT)', 'JORNAL', 'LANCAMENTOS', 'LED', 'MARKETING', 'MERCADO LIVRE', 'MIDIA ON OUTROS', 'MIDIA/ONLINE', 'MKT', 'OUTDOOR', 'PANFLETOS', 'PATROCINIO', 'PORTAL GERACAO', 'PROSPECÇÃO', 'PUBLICIDADE E PROPAGANDA', 'RADIO', 'RD (MKT)', 'REGISTRO SITE', 'SISTEMAS (MKT)', 'SYONET AUTOMOVEIS', 'TELEVISAO', 'VENDAS EXTERNAS', 'VIDEOS', 'VITRINE', 'WISE (MKT)']
            if not any(r[0] == "2" for r in rateios) and not any(r[0] == "2" for r in rateios_aut):
                # print('nao existe centro custo 2')
                if natureza_financeira_value in natureza_financeira_list:
                    time.sleep(2)
                    check_pis_cofins = janela.child_window(class_name="TCheckBox", found_index=0)
                    try:
                        self.wait_until_interactive(check_pis_cofins)
                    except TimeoutError as e:
                        print(str(e))
                        return
                    time.sleep(3)
                    check_pis_cofins.click_input()
                    # time.sleep(2)
                    time.sleep(2)
                    tab_natureza_credito = janela.child_window(title='Natureza Créditos Pis/Cofins', control_type='TabItem')
                    try:
                        self.wait_until_interactive(tab_natureza_credito)
                    except TimeoutError as e:
                        print(str(e))
                        return
                    tab_natureza_credito.click_input()
                    time.sleep(1)
                    if tipo_docto_value == 'NFE':
                        nat_text_nfe = janela.child_window(class_name='TwwDBLookupCombo', found_index=1)
                        try:
                            self.wait_until_interactive(nat_text_nfe)
                        except TimeoutError as e:
                            print(str(e))
                            return
                        nat_text_nfe.click_input()
                        time.sleep(1)
                        pyautogui.typewrite('Aquisicao de bens utilizados como insumo')
                        time.sleep(2)
                        pyautogui.press('tab')
                    elif tipo_docto_value == 'NFS':
                        nat_text_nfs = janela.child_window(class_name='TwwDBLookupCombo', found_index=0)
                        try:
                            self.wait_until_interactive(nat_text_nfs)
                        except TimeoutError as e:
                            print(str(e))
                            return
                        nat_text_nfs.click_input()
                        time.sleep(1)
                        pyautogui.typewrite('Aquisicao de servicos utilizados como insumo')
                        time.sleep(2)
                        pyautogui.press('tab')
        if tipo_docto_value == 'NFS':
            time.sleep(2)
            modelo_fiscal = janela.child_window(title='Modelo Fiscal / Chave e Outros')
            chave_nfse_value = self.get_codigo_from_pdf(num_nf_value, id_solicitacao)
            chave_nfse_sp_tesseract = self.get_codigo_from_pdf(num_nf_value, id_solicitacao)
            try:
                self.wait_until_interactive(modelo_fiscal)
            except TimeoutError as e:
                print(str(e))
                return
            modelo_fiscal.click_input()
            time.sleep(2)
            botao_nfe = janela.child_window(title='NF-e')
            for _ in range(2):
                botao_nfe.click_input()
                time.sleep(0.5)
            for _ in range(2):
                pyautogui.press("tab")
                time.sleep(0.5)
            time.sleep(2)
            if self.xml_existe == False:
                if chave_nfse_value is not None:
                    pyautogui.typewrite(chave_nfse_value)
                else:
                    if chave_nfse_sp_tesseract is not None:
                        pyautogui.typewrite(chave_nfse_sp_tesseract)
            else:
                chave_nfs_value_xml = self.get_verification_code_from_xml_nfs(num_nf_value, id_solicitacao)
                if chave_nfs_value_xml:
                    pyautogui.typewrite(chave_nfs_value_xml)

        time.sleep(3)
        if tipo_docto_value == 'NFE':
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
            valor_nf = self.get_vnf_from_xml(chave_acesso, num_nf_value, id_solicitacao)

            grouped_by_nature = {}
            for item in cfop_results:
                cfop = item['CFOP']
                cst = item.get('CST')  # Busca CST. Se não existir, será None.
                valor_float = float(item['vProd'].replace(',', '.'))
                if cfop == "5929" and cst in ("10", "30", "60", "70", "61"):
                    mapped_value = '1407'
                else:
                    mapped_value = CFOP_MAPPING.get(cfop, default_value)
                if mapped_value in grouped_by_nature:
                    grouped_by_nature[mapped_value] += valor_float
                else:
                    grouped_by_nature[mapped_value] = valor_float
            valor_total_decimal = Decimal(valor_nf)
            grouped_by_decimal = {k: Decimal(v).quantize(Decimal('0.01')) for k, v in grouped_by_nature.items()}
            valor_already_added = Decimal(0)
            for idx, (mapped_value, total) in enumerate(grouped_by_decimal.items()):
                if idx != len(grouped_by_decimal) - 1:
                    valor = total.quantize(Decimal('0.01'))
                    valor_already_added += valor
                else:  # Se for o último item, ajuste para o total da nota
                    valor = valor_total_decimal - valor_already_added

                # valor_str = str(valor).replace('.', ',')
                valor = Decimal(valor).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
                valor_str = str(valor).replace('.',',')
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
        compare_values = self.compare_values_rateio()
        if compare_values == False:
            total_disponivel = self.get_total_disponivel()
            self.inserir_rateio(rateios, conta_contabil, valor_sg, valor_value, usa_rateio_centro_custo, rateios_aut, total_disponivel)
        else:
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
            pyautogui.typewrite('Boleto')
            pyautogui.press('tab')
        elif tipo_pagamento_value in ('P','D'):
            pyautogui.typewrite('Deposito')
            pyautogui.press('tab')
        elif tipo_pagamento_value in ('E', 'C', 'O'):
            pyautogui.typewrite('Especie')
            pyautogui.press('tab')
        try:
            self.wait_until_interactive(natureza_despesa)
        except TimeoutError as e:
            print(str(e))
            return
        natureza_despesa.click_input()
        time.sleep(2)
        pyautogui.typewrite(natureza_financeira_value)
        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(2)
        if tipo_pagamento_value == 'B':
            for boleto in boletos:
                vencimento.click_input()
                time.sleep(1)
                pyautogui.doubleClick()
                time.sleep(1)
                pyautogui.press("backspace")
                print(boleto[0])
                vencimento.type_keys(boleto[0])
                time.sleep(1)
                self.click_specific_button_confidence_9(barra_boleto)
                pyautogui.press("down")
                print("boleto")
        elif tipo_pagamento_value != 'B':
            for n in num_parcelas:
                vencimento.click_input()
                time.sleep(1)
                pyautogui.doubleClick()
                time.sleep(1)
                pyautogui.press("backspace")
                print(n[0])
                vencimento.type_keys(n[0])
                time.sleep(1)
                self.click_specific_button_confidence_9(barra_boleto)
                pyautogui.press("down")
                print("nota")
        print(terceiro)
        if terceiro == 'S':
            self.janela_se_terceiro(numeroos)
        time.sleep(2)
        try:
            if tipo_docto_value == 'NFS' and cod_matriz == '31':
                
                prefeitura = janela.child_window(title='Prefeitura')
                try:
                    self.wait_until_interactive(prefeitura)
                except TimeoutError as e:
                    print(str(e))
                    return
                prefeitura.click_input()
                
                time.sleep(1)
                cnae = janela.child_window(class_name="TBtnWinControl", found_index=2)
                try:
                    self.wait_until_interactive(cnae)
                except TimeoutError as e:
                    print(str(e))
                    return
                cnae.click_input()
                time.sleep(1)
                pyautogui.typewrite('Tributado integralmente')
                time.sleep(1)
                pyautogui.press('enter')

                cfps = janela.child_window(class_name="TBtnWinControl", found_index=1)
                try:
                    self.wait_until_interactive(cfps)
                except TimeoutError as e:
                    print(str(e))
                    return
                cfps.click_input()
                time.sleep(2)
                cidade_empresa = self.consulta_cidade_empresa(empresa)
                if cidade_empresa == cidade_cliente:
                    pyautogui.typewrite('Para Tomador ou destinario estabelecido ou domiciliado no municipio')
                elif cidade_empresa != cidade_cliente and estado == 'SC':
                    pyautogui.typewrite('Para Tomador ou destinario estabelecido ou domiciliado fora do municipio')
                elif estado != 'SC':
                    pyautogui.typewrite('Para Tomador ou destinario estabelecido ou domiciliado em outro estado da Federacao')
                time.sleep(2)
                pyautogui.press('enter')
                time.sleep(1)
                cst = janela.child_window(class_name="TBtnWinControl", found_index=0)
                try:
                    self.wait_until_interactive(cst)
                except TimeoutError as e:
                    print(str(e))
                    return
                cst.click_input()
                time.sleep(1)
                pyautogui.typewrite('Tributado integralmente')
                time.sleep(1)
                pyautogui.press('enter')
        except:
            pass

        submit_button.click_input()
        time.sleep(2)
        if estado != 'SC':
            pyautogui.press('tab')
            pyautogui.press('enter')
        # end_time = time.time() + 30 
        # loop_continue = True
        # while loop_continue and time.time() < end_time:
        #     for elem in janela.children():
        #         if "Aviso NF-e" in elem.window_text():
        #             pyautogui.press("enter")
        #             loop_continue = False
        #             break
        time.sleep(5)
        start_time = time.time()
        try:
            while time.time() - start_time < 30:  
                for elem in janela.children():
                    if "Aviso NF-e" in elem.window_text():
                        pyautogui.press("enter")
                        break
                time.sleep(0.5)  
        except:
            pyautogui.press("enter")

    def click_specific_button_confidence_9(self, button_image_path, confidence_level=0.9, timeout=10):
            end_time = time.time() + timeout
            while time.time() < end_time:
                button_location = pyautogui.locateOnScreen(button_image_path, confidence=confidence_level)
                if button_location:
                    button_x, button_y, button_width, button_height = button_location
                    button_center_x = button_x + button_width // 2
                    button_center_y = button_y + button_height // 2
                    pyautogui.click(button_center_x, button_center_y)
                    print("Botão clicado!")
                    return
                time.sleep(1)  # Espera um segundo antes de tentar novamente
            
            print("Botão não encontrado.")

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
        select_os = r"C:\Users\user\Documents\RPA_Project\imagens\select_os.PNG"
        numero_os = janela.child_window(class_name="TOvcNumericField", found_index=0)
        try:
            self.wait_until_interactive(numero_os)
        except TimeoutError as e:
            print(str(e))
            return
        if numeroos != '':
            relacao_os.click_input()
            time.sleep(1)
            num_os = self.dividir_os_em_blocos(numeroos)
            time.sleep(1)
            if len(numeroos) > 7:
                for num in num_os:
                    numero_os.click_input()
                    time.sleep(1)
                    pyautogui.doubleClick()
                    pyautogui.press('backspace')
                    time.sleep(1)
                    # pyautogui.typewrite(num)
                    numero_os.type_keys(num)
                    time.sleep(1)  
                    self.click_specific_button(pesquisa_os)
                    time.sleep(2)
                    self.click_specific_button(select_os)
            else:
                numero_os.click_input()
                time.sleep(1)
                numero_os.type_keys(numeroos)
                time.sleep(1)  
                self.click_specific_button(pesquisa_os)
                time.sleep(2)
                self.click_specific_button(select_os)
        else:
            relacao_os.click_input()
            time.sleep(2)
            self.click_specific_button(cadeado)
            time.sleep(0.2)
            pyautogui.press('tab')
            time.sleep(0.2)
            pyautogui.press('enter')

    def dividir_os_em_blocos(self, os):
        return os.split(',')


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
            erro_msg = "Falha: Janela de ficha de controle de pagamento não está visível."
            print(erro_msg)
            raise Exception(erro_msg)
            # return
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
        time.sleep(3)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela Ace Viewer")
            return 
        janela = app[title]
        pyautogui.hotkey('alt', 'f')
        time.sleep(2)
        pyautogui.press('s')

    def type_slowly(self, element, text, delay=0.3):
        """Types the text into the element with a delay between each keystroke."""
        for character in text:
            element.type_keys(character)
            time.sleep(delay)

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
        time.sleep(2)
        try:
            get_document = janela.child_window(title="APs (fixo)")
            get_document.click_input()
        except:
            pass
        time.sleep(2)
        file_name = janela.child_window(class_name="Edit")
        file_type = janela.child_window(class_name="AppControlHost", found_index=1)
        # Set Comands
        time.sleep(2)
        try:
            self.wait_until_interactive(file_name)
        except TimeoutError as e:
            print(str(e))
            return
        time.sleep(2)
        # file_name.type_keys(f"AP_{num_docto}{id_solicitacao}")
        self.type_slowly(file_name, f"AP_{num_docto}{id_solicitacao}")
        time.sleep(2)
        pyautogui.press('tab')
        time.sleep(2)
        try:
            self.wait_until_interactive(file_type)
        except TimeoutError as e:
            print(str(e))
            return
        file_type.click_input()
        time.sleep(1)
        for _ in range(2):
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
        for _ in range(2):
            pyautogui.press('enter')
            time.sleep(1)
        # time.sleep(3)
        self.wait_microsoft_edge(num_docto, id_solicitacao)
        # pyautogui.hotkey('ctrl', 'w')


    def wait_microsoft_edge(self, num_docto, id_solicitacao):
        title = f"AP_{num_docto}{id_solicitacao}.pdf - Perfil 1 — Microsoft​ Edge"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela do microsoft edge não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela microsoft edge")
            return 
        pyautogui.hotkey('ctrl', 'w')

    def click_specific_button(self, button_image_path, confidence_level=0.8, timeout=10):
            end_time = time.time() + timeout
            while time.time() < end_time:
                button_location = pyautogui.locateOnScreen(button_image_path, confidence=confidence_level)
                if button_location:
                    button_x, button_y, button_width, button_height = button_location
                    button_center_x = button_x + button_width // 2
                    button_center_y = button_y + button_height // 2
                    pyautogui.click(button_center_x, button_center_y)
                    print("Botão clicado!")
                    return
                time.sleep(1)
            
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
        title = "Sistema Financeiro - SISFIN"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de Sistema Financeiro - SISFIN não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='win32').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a Sistema Financeiro - SISFIN")
            return 
        janela = app[title]
        janela.set_focus()
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
        # try:
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
        # except:
        #     pass
        
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
   
    def inserir_rateio(self, rateios, conta_contabil, valor_sg, vlr_nf, usa_rateio_centro_custo, rateios_aut, total_disponivel=None):
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
                # vlr_rateio = self.get_valor_rateio_manual(rateio[1], valor_sg, vlr_nf)
                vlr_rateio = self.get_valor_rateio_manual(rateio[1], valor_sg, vlr_nf, total_disponivel)
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
                # vlr_rateio = self.get_valor_rateio_aut(vlr_nf, rateio_aut[1])
                vlr_rateio = self.get_valor_rateio_aut(vlr_nf, rateio_aut[1], total_disponivel)
                if indice != tamanho_aut - 1:
                    valor_rateio.type_keys(vlr_rateio)
                time.sleep(1)
                confirmar.click_input()
                time.sleep(1)

    def get_valor_rateio_manual(self, valor_centro_custo, total_sg, valor_nf, total_disponivel=None):
        # valor_nf = valor_nf.replace(',', '.')
        # valor_nf = float(valor_nf) if valor_nf else 0.0  
        # proporcao_centro_custo = valor_centro_custo / total_sg
        # valor_rateio = proporcao_centro_custo * valor_nf
        # valor_rateio_str = "{:.2f}".format(valor_rateio)
        # valor_rateio_str = valor_rateio_str.lstrip('0')
        # return valor_rateio_str.replace('.', ',')
        valor_base = total_disponivel if total_disponivel else valor_nf
        valor_base = valor_base.replace(',', '.')
        valor_base = float(valor_base) if valor_base else 0.0  
        proporcao_centro_custo = valor_centro_custo / total_sg
        valor_rateio = proporcao_centro_custo * valor_base
        valor_rateio_str = "{:.2f}".format(valor_rateio)
        valor_rateio_str = valor_rateio_str.lstrip('0')
        return valor_rateio_str.replace('.', ',')


    def get_valor_rateio_aut(self, valor_nf, porcentagem, total_disponivel=None):
        # valor_nf = valor_nf.replace(',', '.')
        # valor_nf = float(valor_nf) if valor_nf else 0.0  
        # porcentagem_decimal = float(porcentagem) / 100
        # valor_rateio = porcentagem_decimal * valor_nf
        # valor_rateio_str = "{:.2f}".format(valor_rateio)
        # valor_rateio_str = valor_rateio_str.lstrip('0')
        # return valor_rateio_str.replace('.', ',')
        valor_base = total_disponivel if total_disponivel else valor_nf
        valor_base = valor_base.replace(',', '.')
        valor_base = float(valor_base) if valor_base else 0.0  
        porcentagem_decimal = float(porcentagem) / 100
        valor_rateio = porcentagem_decimal * valor_base
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
        if self.xml_existe == False:
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
        if self.xml_existe == False:
            pattern = r'(?<!\d)5102(?!\d)'
            result = re.search(pattern, text)
            return True if result else False

    def get_chave_acesso_pdfplumber(self, num_docto, id_solicitacao):
        if self.xml_existe == 'False':
            caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
            texto_total = ""
            with pdfplumber.open(caminho_pdf) as pdf:
                for pagina in pdf.pages:
                    texto = pagina.extract_text()
                    if texto:
                        texto_total += texto + '\n'

            # Primeira tentativa: Busca por uma sequência de 44 dígitos em blocos de 4
            chave_acesso_match = re.search(r'((\d{4}\s*){10}\d{4})', texto_total)

            # Segunda tentativa: Busca direta por uma sequência de 44 dígitos
            if not chave_acesso_match:
                chave_acesso_match = re.search(r'(\d{44})', texto_total)

            # Terceira tentativa: Busca pelo padrão específico com pontos e hífen
            if not chave_acesso_match:
                chave_acesso_match = re.search(r'\d{2}\.\d{2}\.\d{2}\.\d{14}\.\d{2}\.\d{3}\.\d{9}\.\d{9}-\d', texto_total)

            chave_acesso = ''.join(re.findall(r'\d', chave_acesso_match.group(0))) if chave_acesso_match else None

            return chave_acesso
    
    def get_chave_acesso_tesseract(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
            pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
            caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
            imagens = convert_from_path(caminho_pdf)
            texto_total = ""
            
            for imagem in imagens:
                texto = pytesseract.image_to_string(imagem)
                if texto:
                    texto_total += texto + '\n'
            
            chave_acesso_match = re.search(r'(\d{44})', texto_total)
            if not chave_acesso_match:
                chave_acesso_match = re.search(r'((\d{4}\s*){10}\d{4})', texto_total)
            chave_acesso = chave_acesso_match.group(0) if chave_acesso_match else None
            
            if chave_acesso:
                chave_acesso = ''.join(re.findall(r'\d', chave_acesso))
            
            return chave_acesso

    def get_chave_acesso(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
            chave_pdfplumber = self.get_chave_acesso_pdfplumber(num_docto, id_solicitacao)
            if chave_pdfplumber:
                return chave_pdfplumber
            elif chave_pdfplumber is None:
                return self.get_chave_acesso_tesseract(num_docto, id_solicitacao)

    def get_modelo(self, chave_acesso):
        return chave_acesso[20:22] if len(chave_acesso) > 21 else "Não foi possível extrair os números"
    
    def get_serie(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
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
        else:
            return None
    
    def get_valor_icms(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
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
        
    def check_conditions(self, tipo_docto, natureza, vencimento, inss, irff, piscofinscsl, tipo_pagamento, icms, boletos, cod_nfse, os, iss_value):
        current_date = datetime.date.today()
        converted_vencimento = self.convert_to_date(vencimento)
        # if tipo_docto != "NFE":
        #     return False, "Tipo de documento inválido"
        # if tipo_docto == 'NFE':
        #     if not chave_acesso and self.xml_existe == False:
        #         return False, "chave de acesso de produto não encontrada"
        # if tipo_docto == 'NFS':
        #     if not cod_nfse and self.xml_existe == False:
        #         return False, "chave de acesso de servico não encontrada"
        # if tipo_pagamento == 'B':
        #     for boleto in boletos:
        #         converted_boleto_date = self.convert_to_date(boleto[0])
        #         if converted_boleto_date < current_date:
        #             return False, "O primeiro boleto tem vencimento anterior à data atual"
        #         break 
        # elif tipo_pagamento != 'B':
        #     if converted_vencimento < current_date:
        #         return False, "vencimento menor que data de efetivacao"
        # if inss:
        #     return False, "inss encontrado"
        # if irff:
        #     return False, "irff encontrado"
        # if iss_value:
        #     return False, "iss encontrado"
        # if piscofinscsl:
        #     return False, "piscofinscsl encontrado"
        if tipo_pagamento not in ('B', 'A', 'P', 'D', 'E', 'C', 'O'):
            return False, "tipo de pagamento diferente do configurado"
        if os != '':
            if len(os) > 7 and ',' not in os:
                return False, "Os não estão separadas por vírgula"
        
        
        return True, "Condições aceitas"

    def check_cod_nfs(self, tipo_docto, cod_nfse):
        if tipo_docto == 'NFS':
            if not cod_nfse and self.xml_existe == False:
                return False
                    
        return True
        
    def check_tribut(self, inss, irff, piscofinscsl, iss_value):
        if inss:
            return False
        if irff:
            return False
        if iss_value:
            return False
        if piscofinscsl:
            return False
        
        return True

    def send_message_pre_verification(self, msg, id_solicitacao, numerodocto):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        # base_url = f"https://api.telegram.org/bot6563290849:AAEVdZmaKFWOTmUdIYeAcGhq3_cUu4sX2qE/sendMessage"
        base_url = f"https://api.telegram.org/bot{token}/sendMessage"
        for chat_id in chat_ids:
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
        chat_ids_results = database.consultar_chat_id_dev()
        token_result = database.consultar_token_bot()
        chat_ids = [result[0] for result in chat_ids_results]
        token = token_result[0][0]
        error_message = traceback.format_exc()
        detailed_msg = f"[TRACEBACK ERROR]: {error_message}, id_solicitacao: {id_solicitacao}, numero_docto: {numerodocto}"
        
        base_url = f"https://api.telegram.org/bot{token}/sendMessage"

        for chat_id in chat_ids:
            payload = {
                'chat_id': chat_id,
                'text': detailed_msg
            }
            response = requests.post(base_url, data=payload)
            if response.status_code != 200:
                print(f"Failed to send traceback message to chat_id {chat_id}. Response: {response.content}")
            else:
                print(f"Traceback message sent successfully to chat_id {chat_id} on Telegram!")

    def send_success_message(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}"
        
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

    def send_success_message_telecomunication(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}, obs: Telecomunicacao"
        
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


    def send_success_message_not_serie(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}, obs: serie nao existente, verificar"
        
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

    def send_success_message_compromisso(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}, obs: Compromisso"
        
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

    def send_success_message_fatura(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}, obs: Fatura"
        
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

    def send_success_message_cod_nfs_exception(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}, obs: sem codigo de acesso de nfs"
        
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

    def send_success_message_cod_tribut_exceptions(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}, obs: retencao encontrada"
        
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

    def send_success_message_cod_tribut_cod_nfs_exceptions(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Lançamento efetuado com sucesso, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}, obs: retencao encontrada e chave de acesso de servico nao encontrado"
        
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

    def send_ap_nao_existe_message(self, id_solicitacao, numero_nota, tipo_nota):
        chat_ids_results = database.consultar_chat_id_dev()
        token_result = database.consultar_token_bot()

        # chat_id1, chat_id2 = chat_ids_results[0][0], chat_ids_results[1][0]
        # chat_ids = [chat_id1, chat_id2]
        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"AP nao existe, id solicitacao: {id_solicitacao}, numero da nota: {numero_nota}, tipo da nota: {tipo_nota}"
        
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

    def send_failure_try_message(self, id_solicitacao, numero_nota):
        chat_ids_results = database.consultar_chat_id()
        token_result = database.consultar_token_bot()

        chat_ids = [result[0] for result in chat_ids_results]

        token = token_result[0][0]
        success_msg = f"Solicitacao Rejeitada por erro, id_solicitacao: {id_solicitacao}, numero da nota: {numero_nota}"
        
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

    def get_vnf_from_xml(self, chave_acesso, num_docto, id_solicitacao):
        if self.xml_existe == False:
            path = rf'C:\Users\user\Downloads\{chave_acesso}.xml'
        else:
            path = rf'C:\Users\user\Downloads\xml_{num_docto}_{id_solicitacao}.xml'
        tree = ET.parse(path)
        root = tree.getroot()

        # Identificando o namespace
        ns = None
        for elem in root.iter():
            if '}' in elem.tag:
                ns = {'ns': elem.tag.split('}')[0].strip('{')}
                break
        if ns is None:
            raise ValueError("Namespace não encontrado no XML.")

        # Buscando o elemento vNF
        total_nf_elem = root.find('.//ns:vNF', namespaces=ns)
        if total_nf_elem is not None:
            valor_total_nf = float(total_nf_elem.text.replace(',', '.'))
            return valor_total_nf
        else:
            raise ValueError("Elemento vNF não encontrado no XML.")

    def get_data_from_xml(self, chave_acesso, num_docto, id_solicitacao):
        if self.xml_existe == False:
            path = rf'C:\Users\user\Downloads\{chave_acesso}.xml'
        else:
            path = rf'C:\Users\user\Downloads\xml_{num_docto}_{id_solicitacao}.xml'
        tree = ET.parse(path)
        root = tree.getroot()

        ns = None
        for elem in root.iter():
            if '}' in elem.tag:
                ns = {'ns': elem.tag.split('}')[0].strip('{')}
                break
        if ns is None:
            raise ValueError("Namespace não encontrado no XML.")
        
        cfop_values = {}
        det_elements = root.findall('.//ns:det', namespaces=ns)
        for det in det_elements:
            cfop_elem = det.find('.//ns:CFOP', namespaces=ns)
            vprod_elem = det.find('.//ns:vProd', namespaces=ns)
            vdesc_elem = det.find('.//ns:vDesc', namespaces=ns)
            
            icms_element = det.find('.//ns:ICMS', namespaces=ns)
            cst_elem = None
            if icms_element is not None:
                for child in icms_element:
                    cst_elem = child.find('.//ns:CST', namespaces=ns)
                    if cst_elem is not None:
                        break
            cst = cst_elem.text if cst_elem is not None else None

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
        results = [{'CFOP': key, 'vProd': str(value).replace('.', ','), 'CST': cst} for key, value in cfop_values.items()]
        return results

    def get_data_emissao_from_xml(self, chave_acesso, num_docto, id_solicitacao):
        if self.xml_existe == False:
            path = rf'C:\Users\user\Downloads\{chave_acesso}.xml'
        else:
            path = rf'C:\Users\user\Downloads\xml_{num_docto}_{id_solicitacao}.xml'
        tree = ET.parse(path)
        root = tree.getroot()
        ns = None
        for elem in root.iter():
            if '}' in elem.tag:
                ns = {'ns': elem.tag.split('}')[0].strip('{')}
                break
        if ns is None:
            raise ValueError("Namespace não encontrado no XML.")
        dhemi_element = root.find('.//ns:dhEmi', namespaces=ns)
        if dhemi_element is not None:
            date_string = dhemi_element.text
            date_part = date_string.split('T')[0]
            date_object = datetime.datetime.strptime(date_part, "%Y-%m-%d")
            formatted_date = date_object.strftime('%d-%m-%Y')
            
            return formatted_date

    def remover_arquivo_nota(self, numerodoct, id_solicitacao):
        nome_arquivo = f"nota_{numerodoct}{id_solicitacao}.pdf"
        caminho_arquivo = os.path.join('C:\\Users\\user\\Documents\\APs', nome_arquivo)
    
        if os.path.exists(caminho_arquivo):
            try:
                os.remove(caminho_arquivo)
                print(f"Arquivo {nome_arquivo} removido com sucesso!")
            except Exception as e:
                print(f"Erro ao tentar remover o arquivo {nome_arquivo}. Erro: {e}")
        else:
            print(f"Arquivo {nome_arquivo} não existe.")

    def remover_arquivo_xml(self, numerodoct, id_solicitacao):
        nome_arquivo = f"xml_{numerodoct}_{id_solicitacao}.xml"
        caminho_arquivo = os.path.join('C:\\Users\\user\\Downloads', nome_arquivo)
    
        if os.path.exists(caminho_arquivo):
            try:
                os.remove(caminho_arquivo)
                print(f"Arquivo {nome_arquivo} removido com sucesso!")
            except Exception as e:
                print(f"Erro ao tentar remover o arquivo {nome_arquivo}. Erro: {e}")
        else:
            print(f"Arquivo {nome_arquivo} não existe.")

    # def remover_arquivo(self, chave_acesso, num_nota, num_solicitacao):
    #     if self.xml_existe == False:
    #         nome_arquivo = f"{chave_acesso}.xml"
    #         caminho_arquivo = os.path.join('C:\\Users\\user\\Downloads', nome_arquivo)
    #     else:
    #         nome_arquivo = f"xml_{num_nota}_{num_solicitacao}.xml"
    #         caminho_arquivo = os.path.join('C:\\Users\\user\\Downloads', nome_arquivo)
            
    #     if os.path.exists(caminho_arquivo):
    #         try:
    #             os.remove(caminho_arquivo)
    #             print(f"Arquivo {nome_arquivo} removido com sucesso!")
    #         except Exception as e:
    #             print(f"Erro ao tentar remover o arquivo {nome_arquivo}. Erro: {e}")
    #     else:
    #         print(f"Arquivo {nome_arquivo} não existe.")

    def compare_values_rateio(self):
        title = "Entrada Diversas / Operação: 52-Entrada Diversas"
        if not self.esperar_janela_visivel(title, timeout=60):
            print("Falha: Janela de Entrada Diversas / Operação: 52-Entrada Diversas não está visível.")
            return
        time.sleep(2)
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de Entrada Diversas / Operação: 52-Entrada Diversas")
            return False  
        janela = app[title]
        disponivel = janela.child_window(class_name="TOvcPictureField", found_index=3)
        valor_disponivel = disponivel.legacy_properties()['Name'].replace(" ", "")
        v_total_control = janela.child_window(class_name="TOvcDbPictureField", found_index=0)
        valor_total = v_total_control.legacy_properties()['Name'].replace(" ", "")
        
        return valor_disponivel == valor_total


    def get_serie_from_xml(self, chave_acesso, num_docto, id_solicitacao):
        if self.xml_existe == False:
            path = rf'C:\Users\user\Downloads\{chave_acesso}.xml'
        else:
            path = rf'C:\Users\user\Downloads\xml_{num_docto}_{id_solicitacao}.xml'
        tree = ET.parse(path)
        root = tree.getroot()
        ns = None
        for elem in root.iter():
            if '}' in elem.tag:
                ns = {'ns': elem.tag.split('}')[0].strip('{')}
                break
        if ns is None:
            raise ValueError("Namespace não encontrado no XML.")
        
        serie_elem = root.find('.//ns:serie', namespaces=ns)
        serie = serie_elem.text if serie_elem is not None else None
        
        return serie

    def get_codigo_from_pdf(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
            try:
                texto_total = self.extract_text_from_pdf_with_pdfplumber(num_docto, id_solicitacao)
            except:
                try:
                    texto_total = self.extract_text_from_pdf(num_docto, id_solicitacao)
                except:
                    return None
            # print(texto_total)

            # Padrão para "Chave de Acesso da NFS-e"
            chave_acesso_pattern = r'ChavedeAcessodaNFS-e\s*(\d{8,})'
            chave_acesso_match = re.search(chave_acesso_pattern, texto_total)
            if chave_acesso_match:
                return chave_acesso_match.group(1)

            # Padrão para diversos códigos
            codigos_pattern1 = (r'(?i)(Código\s*de\s*Verificação|Código\s*de\s*Validação|Autenticidade|Chave\s*Acesso|'
                                r'Chave\s*de\s*Acesso|Chave\s*de\s*Acesso\s*daNFS-e)[^A-Z0-9]*?'
                                r'(?:[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*)?(\d{8,}\b)(?!.*000000)')
            codigo_match1 = re.search(codigos_pattern1, texto_total)
            if codigo_match1 and "000000" not in codigo_match1.group(2):
                return codigo_match1.group(2)

            # Padrão para "Código de autenticação da NFSe"
            codigos_pattern5 = (r'(?i)Código de autenticação da NFSe:\s*([^\s]+)')
            codigo_match5 = re.search(codigos_pattern5, texto_total)
            if codigo_match5:
                return codigo_match5.group(1)

            # Padrão para "Autenticidade"
            autenticidade_pattern = r'Autenticidade\s*(?:[\s\S]*?)(\d{10,})'
            autenticidade_match = re.search(autenticidade_pattern, texto_total)
            if autenticidade_match:
                return autenticidade_match.group(1)

            # Padrão para "Identificador"
            codigos_pattern4 = (r'(?i)Identificador[\s\S]*?((?:\d{4}\s){9}\d{4})')
            codigo_match4 = re.search(codigos_pattern4, texto_total)
            if codigo_match4:
                codigo4 = codigo_match4.group(1).replace(' ', '')  # remove os espaços do código capturado
                if "000000" not in codigo4:
                    return codigo4
            
            if not codigo_match5 and not codigo_match4 and not codigo_match1 and not chave_acesso_match:
                long_number_pattern = r'\b\d{37,}\b'
                long_number_match = re.search(long_number_pattern, texto_total)
                if long_number_match:
                    return long_number_match.group()

            return None  # Retorna None se nenhum código for encontrado

    # def get_codigo_from_pdf(self, num_docto, id_solicitacao):
    #     if self.xml_existe == False:
    #         texto_total = self.extract_text_from_pdf_with_pdfplumber(num_docto, id_solicitacao)
    #         # print(texto_total)

    #         # Novo padrão para "Chave de Acesso da NFS-e"
    #         chave_acesso_pattern = r'ChavedeAcessodaNFS-e\s*(\d{8,})'
    #         chave_acesso_match = re.search(chave_acesso_pattern, texto_total)
    #         if chave_acesso_match:
    #             return chave_acesso_match.group(1)

    #         codigos_pattern1 = (r'(?i)(Código\s*de\s*Verificação|Código\s*de\s*Validação|Autenticidade|Chave\s*Acesso|'
    #                             r'Chave\s*de\s*Acesso|Chave\s*de\s*Acesso\s*daNFS-e)[^A-Z0-9]*?'
    #                             r'(?:[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*[A-Z]*\s*)?(\d{8,}\b)(?!.*000000)')
    #         codigo_match1 = re.search(codigos_pattern1, texto_total)
    #         codigo1 = codigo_match1.group(2) if codigo_match1 and "000000" not in codigo_match1.group(2) else None

    #         if not codigo1:
    #             codigos_pattern5 = (r'(?i)Código de autenticação da NFSe:\s*([^\s]+)')
    #             codigo_match5 = re.search(codigos_pattern5, texto_total)
    #             if codigo_match5:
    #                 codigo5 = codigo_match5.group(1)
    #                 return codigo5

    #         if not codigo1:
    #             codigos_pattern4 = (r'(?i)Identificador[\s\S]*?((?:\d{4}\s){9}\d{4})')
    #             codigo_match4 = re.search(codigos_pattern4, texto_total)
    #             if codigo_match4:
    #                 codigo4 = codigo_match4.group(1).replace(' ', '')  # remove os espaços do código capturado
    #                 if "000000" not in codigo4:
    #                     return codigo4

    #         return None  # Retorna None se nenhum código for encontrado
    #     else:
    #         return None
        
    def get_service_code(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
            texto_total = self.extract_text_from_pdf_with_pdfplumber(num_docto, id_solicitacao) 
            # print(texto_total) 
            
            pattern1 = r'C[oó]digodeTributa[cç][aã]oNacional\s*[^0-9]*([1-9]\d{0,1}\.\d{1,2})'
            pattern2 = r'(c[oó]dig(?:o|os)?\s*(?:de|dos)?\s*servi[cç][oó]s|atividade\s*do\s*munic[ií]pio|c[oó]digo\s*de\s*tributa[cç][aã]o\s*nacional)\s*[^0-9]*\s*(\d{1,2}\.\d{1,2}(?:\.\d{1,2})?)'
            pattern3 = r'Atividade:\s*[^0-9]*(\d{1,5})'  # Procura pelo termo "Atividade:" seguido por qualquer número de caracteres não numéricos e então um número com até 5 dígitos.
            
            match1 = re.search(pattern1, texto_total, re.IGNORECASE)
            
            if match1:
                code = match1.group(1).replace('.', '')
                return code[:4]
            
            match2 = re.search(pattern2, texto_total, re.IGNORECASE)
            
            if match2:
                code = match2.group(1).replace('.', '')
                return code[:4]
            
            match3 = re.search(pattern3, texto_total, re.IGNORECASE)
            
            if match3:
                return match3.group(1)[:4]

            return None
    
    def consulta_cidade_empresa(self, empresa: str) -> str:
        city_map = {
            "FLORIPA": "FLORIANOPOLIS",
            "TUBARAO": "TUBARAO",
            "LAGES": "LAGES",
            "CRICIUMA": "CRICIUMA",
        }

        city = empresa.split()[-1].upper()
        return city_map.get(city, city)

    def get_aliquota(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
            texto_total = self.extract_text_from_pdf_with_pdfplumber(num_docto, id_solicitacao) 
            # print(texto_total)
            aliquota_value = None
            # Lages
            # Tentativa 1: Após "Base de cálculo (%)", pegue o número após o 'x'
            try:
                pattern1 = r'Base de cálculo \(%\).*?\d+(?:,\d{1,2})?x(\d+(?:,\d{1,2})?)='
                match = re.search(pattern1, texto_total, re.IGNORECASE | re.DOTALL)
                if match:
                    aliquota_value = match.group(1)
            except:
                pass

            # Criciuma
            # Tentativa 2: Após "Alíquota ISS", pegue o primeiro valor com % e retorne as duas primeiras casas após a vírgula
            if not aliquota_value:
                try:
                    pattern2 = r'Alíquota ISS.*?(\d+,\d{1,4})\s?%'
                    match = re.search(pattern2, texto_total, re.IGNORECASE | re.DOTALL)
                    if match:
                        full_value = match.group(1)
                        main, decimal = full_value.split(',')
                        decimal = decimal[:4] # limitar em 4 casas decimais
                        aliquota_value = f"{main},{decimal}"
                except:
                    pass
            # Apos aliquota ISS, pegue o primeiro valor que tem mais de 4 casas decimais apos a virgula e formate para 2 casas
            if not aliquota_value:
                try:
                    pattern5 = r'Alíquota\s+ISS.*?([\d\.]{1,},\d{5,})'
                    match = re.search(pattern5, texto_total, re.IGNORECASE | re.DOTALL)
                    if match:
                        raw_value = match.group(1)
                        # Convertendo a string para um float e formatando
                        formatted_value = "{:.2f}".format(float(raw_value.replace(',', '.'))).replace('.', ',')
                        aliquota_value = formatted_value
                        if aliquota_value != "0,00":
                            return aliquota_value
                except:
                    pass

                # Tentativa 4: Captura um valor após a string "Aliquota=" e transforma em formato 0,00
            if not aliquota_value:
                try:
                    pattern4 = r'Aliquota=(\d+)'  # este regex irá pegar um ou mais dígitos após "Aliquota="
                    match = re.search(pattern4, texto_total)
                    if match:
                        aliquota_int = match.group(1)
                        aliquota_value = "{:.2f}".format(float(aliquota_int)).replace(".", ",")
                        if aliquota_value != "0,00":
                            return aliquota_value
                except Exception as e:
                    print(f"Erro ao processar Aliquota: {e}")

            if not aliquota_value:
                try:
                    pattern6 = r'Alíquota\s+\(%\).*?([\d\.]{1,},\d{3,})'
                    match = re.search(pattern6, texto_total, re.IGNORECASE | re.DOTALL)
                    if match:
                        raw_value = match.group(1)
                        # Convertendo a string para um float e formatando
                        formatted_value = "{:.2f}".format(float(raw_value.replace(',', '.'))).replace('.', ',')
                        aliquota_value = formatted_value
                        if aliquota_value != "0,00":
                            return aliquota_value
                        
                except:
                    pass
            
            if not aliquota_value:
                try:
                    pattern7 = r'Base de Cálculo.*?\n.*?(\d+,\d{3,})%'
                    match = re.search(pattern7, texto_total, re.IGNORECASE | re.DOTALL)
                    if match:
                        full_value = match.group(1)
                        main, decimal = full_value.split(',')
                        decimal = decimal[:4]  # Limitar para 4 casas decimais
                        formatted_value = f"{main},{decimal}"
                        return formatted_value
                except:
                    pass

            # Após "Alíquota", pegue o primeiro número que tem 3 ou mais casas após a vírgula
            if not aliquota_value:
                try:
                    # pattern8 = r'Alíquota.*?(\b\d\.\d{3,4}\b)'
                    pattern8 = r'Alíquota\s*(\b\d\.\d{3,4}\b)(?!\s*/)'
                    match = re.search(pattern8, texto_total, re.IGNORECASE | re.DOTALL)
                    if match:
                        aliquota_value = match.group(1).replace('.', ',')  # Substitua o ponto por uma vírgula
                except:
                    pass

            if not aliquota_value:
                try:
                    pattern9 = r'Alíquota \(%\).*?(\d+,\d{1,4})\s?'
                    match = re.search(pattern9, texto_total, re.IGNORECASE | re.DOTALL)

                    if match:
                        aliquota_value = match.group(1)
                        return aliquota_value # Deve imprimir 2,00
                except:
                    pass

            if not aliquota_value:
                try:
                    pattern10 = r'Alíquota\s+Situação[\s\S]*?(\d+%)'
                    match = re.search(pattern10, texto_total, re.IGNORECASE | re.DOTALL)

                    if match:
                        aliquota_value = match.group(1)
                        return aliquota_value 
                except:
                    pass

            if not aliquota_value:
                try:
                    pattern11 = r'Aliquota da Atividade:\s*([\d,]+)'
                    match = re.search(pattern11, texto_total, re.IGNORECASE)

                    if match:
                        aliquota_value = match.group(1)
                        return aliquota_value
                except:
                    pass

            # Brasilia
            # Tentativa 3: Pegar o primeiro valor numérico que tem tanto uma vírgula `,` quanto um ponto `.`
            if not aliquota_value:
                try:
                    pattern3 = r'Alíquota.*?(\d{1,2},\d{1,2}\.)'
                    match = re.search(pattern3, texto_total, re.DOTALL)
                    if match:
                        # Remover o ponto ao final
                        aliquota_value = match.group(1)[:-1]
                except:
                    pass

            if aliquota_value and aliquota_value != "0,00":
                return aliquota_value
            
            return None


    def extract_text_from_pdf_with_pdfplumber(self, num_docto, id_solicitacao):
        if self.xml_existe == False:
            caminho_pdf = rf"C:\Users\user\Documents\APs\nota_{num_docto}{id_solicitacao}.pdf"
            
            texto_total = ""
            try:
                with pdfplumber.open(caminho_pdf) as pdf:
                    for pagina in pdf.pages:
                        texto = pagina.extract_text()
                        if texto:
                            texto_total += texto + '\n'
            except Exception as e:
                print(f"Erro ao ler o PDF: {e}")
                return None
            
            return texto_total
    
    def extract_access_key_from_image(self, num_docto, id_solicitacao):
        texto_total = self.extract_text_from_pdf(num_docto, id_solicitacao)
        # print(texto_total)
        
        try:  
            # pattern = r'- NOTA CARIOCA -\s*([\w-]+)'
            pattern = r'\b[A-Za-z0-9]{4}-[A-Za-z0-9]{4}\b'
            match = re.search(pattern, texto_total)
            if match:
                return match.group()
            else:
                return None
        except:
            pass

    def extract_aliquota_from_image(self, num_docto, id_solicitacao):
        texto_total = self.extract_text_from_pdf(num_docto, id_solicitacao)
        # print(texto_total)
        
        if re.search(r'CARIOCA|MUNICIPIO DE SAO PAULO', texto_total, re.IGNORECASE):
            # sp/rj
            try:
                pattern = r'(\d{1,2},\d{2}\%)'
                match = re.search(pattern, texto_total)
                if match is not None:
                    aliquota = match.group(1)
                    if aliquota == '0,00%':
                        return None
                    return aliquota
            except:
                pass
        return None

    def extract_service_code_from_image(self, num_docto, id_solicitacao):
        texto_total = self.extract_text_from_pdf(num_docto, id_solicitacao)
        # print(texto_total)
        # Primeira tentativa: Servico Prestado
        try:
            pattern = r'(?:Servico Prestado)[\s\S]*?([\d\.]{1,7})'
            match = re.search(pattern, texto_total)
            if match:
                # Remove pontos e considera apenas os primeiros 4 dígitos
                code = ''.join([char for char in match.group(1) if char.isdigit()])[:4]
                # Remove zeros à esquerda
                return code.lstrip('0')
        except Exception as e:
            print("Erro na primeira tentativa:", e)
            
        # Segunda tentativa: Cadigo do Servico
        try:
            pattern = r'(?:Cadigo do Servico)[\s\S]*?([\d\.]{1,7})'
            match = re.search(pattern, texto_total)
            if match:
                # Remove pontos e considera apenas os primeiros 4 dígitos
                code = ''.join([char for char in match.group(1) if char.isdigit()])[:4]
                # Remove zeros à esquerda
                return code.lstrip('0')
        except Exception as e:
            print("Erro na segunda tentativa:", e)
        
        # Se nenhuma das tentativas funcionar, retorne None
        return None

    def get_total_disponivel(self):
        title = "Entrada Diversas / Operação: 52-Entrada Diversas"
        try:
            app = Application(backend='uia').connect(title=title)
        except ElementNotFoundError:
            print("Não foi possível se conectar com a janela de cadastro de nf")
            return 
        janela = app[title]

        # for index in range(0,100):
        total_disponivel = janela.child_window(class_name='TOvcPictureField', found_index=3)
        total_disponivel.click_input()
        time.sleep(0.5)
        pyautogui.doubleClick()
        time.sleep(0.5)
        pyautogui.hotkey('ctrl','c')
        time.sleep(0.5)
        texto_copiado = pyperclip.paste()
        texto_copiado = texto_copiado.strip()

        return texto_copiado
    
    
    # def get_verification_code_from_xml_nfs(self, num_nota, id_solicitacao): 
    #     path = rf'C:\Users\user\Downloads\xml_{num_nota}_{id_solicitacao}.xml'
        
    #     try:
    #         # Tentativa de encontrar 'CodigoVerificacao'
    #         tree = ET.parse(path)
    #         root = tree.getroot()
    #         verification_code_elem = root.find('.//CodigoVerificacao')
    #         if verification_code_elem is not None and verification_code_elem.text:
    #             return verification_code_elem.text
    #     except ET.ParseError as e:
    #         print(f"Erro ao analisar 'CodigoVerificacao': {e}")
    #     except Exception as e:
    #         print(f"Erro desconhecido ao processar 'CodigoVerificacao': {e}")

    #     try:
    #         # Tentativa de extrair o 'Id' do 'infNFSe' com namespace
    #         tree = ET.parse(path)
    #         root = tree.getroot()
    #         namespaces = {'ns': 'http://www.sped.fazenda.gov.br/nfse'}
    #         infNFSe_elem = root.find('.//ns:infNFSe', namespaces)
    #         if infNFSe_elem is not None and 'Id' in infNFSe_elem.attrib:
    #             id_value = infNFSe_elem.attrib['Id']
    #             numeric_id = ''.join(filter(str.isdigit, id_value))
    #             return numeric_id if numeric_id else None
    #     except ET.ParseError as e:
    #         print(f"Erro ao analisar 'infNFSe': {e}")
    #     except Exception as e:
    #         print(f"Erro desconhecido ao processar 'infNFSe': {e}")

    #     try:
    #         # Tentativa de encontrar 'cod_verificador_autenticidade'
    #         tree = ET.parse(path)
    #         root = tree.getroot()
    #         verification_code_element = root.find('.//cod_verificador_autenticidade')
    #         if verification_code_element is not None and verification_code_element.text:
    #             return verification_code_element.text
    #     except ET.ParseError as e:
    #         print(f"Erro ao analisar 'cod_verificador_autenticidade': {e}")
    #     except Exception as e:
    #         print(f"Erro desconhecido ao processar 'cod_verificador_autenticidade': {e}")

    #     return None


    # def get_aliquota_from_xml_nfs(self, num_nota, id_solicitacao):
    #     path = rf'C:\Users\user\Downloads\xml_{num_nota}_{id_solicitacao}.xml'
    #     try:
    #         tree = ET.parse(path)
    #         root = tree.getroot()
            
    #         aliquota_elem = root.find('.//Aliquota')
            
    #         if aliquota_elem is not None and aliquota_elem.text:
    #             # Substitui o ponto por vírgula
    #             return aliquota_elem.text.replace('.', ',')
    #     except: 
    #         print("aliquota nao encontrada com o campo .//Aliquota")
    #     try:
    #         tree = ET.parse(path)
    #         root = tree.getroot()
            
    #         aliquota_element = root.find('.//aliquota_item_lista_servico')
            
    #         if aliquota_element is not None and aliquota_element.text:
    #             # Substitui o ponto por vírgula
    #             return aliquota_element.text.replace('.', ',')
    #     except: 
    #         print("aliquota nao encontrada com o campo .//aliquota_item_lista_servico")
    #     return None

    def get_verification_code_from_xml_nfs(self, num_nota, id_solicitacao): 
        path = rf'C:\Users\user\Downloads\xml_{num_nota}_{id_solicitacao}.xml'
        
        try:
            # Tentativa de encontrar 'CodigoVerificacao'
            tree = ET.parse(path)
            root = tree.getroot()
            verification_code_elem = root.find('.//CodigoVerificacao')
            if verification_code_elem is not None and verification_code_elem.text:
                return verification_code_elem.text
        except ET.ParseError as e:
            print(f"Erro ao analisar 'CodigoVerificacao': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'CodigoVerificacao': {e}")

        try:
            # Tentativa de extrair o 'Id' do 'infNFSe' com namespace
            tree = ET.parse(path)
            root = tree.getroot()
            namespaces = {'ns': 'http://www.sped.fazenda.gov.br/nfse'}
            infNFSe_elem = root.find('.//ns:infNFSe', namespaces)
            if infNFSe_elem is not None and 'Id' in infNFSe_elem.attrib:
                id_value = infNFSe_elem.attrib['Id']
                numeric_id = ''.join(filter(str.isdigit, id_value))
                return numeric_id if numeric_id else None
        except ET.ParseError as e:
            print(f"Erro ao analisar 'infNFSe': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'infNFSe': {e}")

        try:
            # Tentativa de encontrar 'cod_verificador_autenticidade'
            tree = ET.parse(path)
            root = tree.getroot()
            verification_code_element = root.find('.//cod_verificador_autenticidade')
            if verification_code_element is not None and verification_code_element.text:
                return verification_code_element.text
        except ET.ParseError as e:
            print(f"Erro ao analisar 'cod_verificador_autenticidade': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'cod_verificador_autenticidade': {e}")

        try:
            # Tentativa de encontrar 'codigo_autenticidade'
            tree = ET.parse(path)
            root = tree.getroot()
            verification_code_elem = root.find('.//codigo_autenticidade')
            if verification_code_elem is not None and verification_code_elem.text:
                return verification_code_elem.text
        except ET.ParseError as e:
            print(f"Erro ao analisar 'codigo_autenticidade': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'codigo_autenticidade': {e}")

        try:
            # Tentativa de encontrar 'codigoVerificacao' diretamente
            tree = ET.parse(path)
            root = tree.getroot()
            verification_code_elem = root.find('.//codigoVerificacao')
            if verification_code_elem is not None and verification_code_elem.text:
                return verification_code_elem.text
        except ET.ParseError as e:
            print(f"Erro ao analisar 'codigoVerificacao': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'codigoVerificacao': {e}")

        try:
            # Carregar o XML
            tree = ET.parse(path)
            root = tree.getroot()

            # Namespace padrão do XML
            namespaces = {'ns': 'http://www.abrasf.org.br/nfse.xsd'}

            # Encontrar o elemento 'CodigoVerificacao' utilizando o namespace
            verification_code_elem = root.find('.//ns:CodigoVerificacao', namespaces)
            if verification_code_elem is not None:
                return verification_code_elem.text
        except ET.ParseError as e:
            print(f"Erro ao analisar XML: {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar XML: {e}")
        try:
            # Carregar o XML
            tree = ET.parse(path)
            root = tree.getroot()

            # Encontrar o elemento 'CodigoVerificacao' utilizando o namespace específico
            verification_code_elem = root.find('.//{http://www.betha.com.br/e-nota-contribuinte-ws}CodigoVerificacao')
            if verification_code_elem is not None and verification_code_elem.text:
                return verification_code_elem.text
        except ET.ParseError as e:
            print(f"Erro ao analisar 'CodigoVerificacao': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'CodigoVerificacao': {e}")
        try:
            # Carregar o XML
            tree = ET.parse(path)
            root = tree.getroot()

            # Encontrar o elemento 'CodigoVerificacao' ou 'CodigoValidacao'
            verification_code_elem = root.find('.//CodigoVerificacao')
            if verification_code_elem is None or not verification_code_elem.text:
                verification_code_elem = root.find('.//CodigoValidacao')
            if verification_code_elem is not None and verification_code_elem.text:
                return verification_code_elem.text
        except ET.ParseError as e:
            print(f"Erro ao analisar 'CodigoVerificacao' ou 'CodigoValidacao': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'CodigoVerificacao' ou 'CodigoValidacao': {e}")


        return None



    def get_aliquota_from_xml_nfs(self, num_nota, id_solicitacao):
        path = rf'C:\Users\user\Downloads\xml_{num_nota}_{id_solicitacao}.xml'
        try:
            tree = ET.parse(path)
            root = tree.getroot()
            
            aliquota_elem = root.find('.//Aliquota')
            
            if aliquota_elem is not None and aliquota_elem.text:
                # Substitui o ponto por vírgula
                return aliquota_elem.text.replace('.', ',')
        except: 
            print("aliquota nao encontrada com o campo .//Aliquota")
        try:
            tree = ET.parse(path)
            root = tree.getroot()
            
            aliquota_element = root.find('.//aliquota_item_lista_servico')
            
            if aliquota_element is not None and aliquota_element.text:
                # Substitui o ponto por vírgula
                return aliquota_element.text.replace('.', ',')
        except: 
            print("aliquota nao encontrada com o campo .//aliquota_item_lista_servico")
        try:
            tree = ET.parse(path)
            root = tree.getroot()
            
            # Encontrar o elemento <itemServico>
            item_servico_elem = root.find('.//itemServico')
            
            if item_servico_elem is not None:
                # Dentro do elemento <itemServico>, encontrar o elemento <aliquota>
                aliquota_elem = item_servico_elem.find('.//aliquota')
                
                if aliquota_elem is not None and aliquota_elem.text:
                    # Substituir o ponto por vírgula
                    aliquota_value = float(aliquota_elem.text) * 100
                    aliquota_value = str(aliquota_value)
                    aliquota_value = aliquota_value.replace('.', ',')
                        
                    # Se o valor for zero, retornar None
                    if aliquota_value == '0':
                        return None
                    else:
                        return aliquota_value
        except ET.ParseError as e:
            print(f"Erro ao analisar XML: {e}")
        try:
            # Carregar o XML
            tree = ET.parse(path)
            root = tree.getroot()

            # Namespace padrão do XML
            namespaces = {'ns': 'http://www.abrasf.org.br/nfse.xsd'}

            # Encontrar o elemento 'Aliquota' utilizando o namespace
            aliquota_elem = root.find('.//ns:Aliquota', namespaces)
            if aliquota_elem is not None and aliquota_elem.text:
                # Substitui o ponto por vírgula
                return aliquota_elem.text.replace('.', ',')
        except ET.ParseError as e:
            print(f"Erro ao analisar XML: {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar XML: {e}")
        try:
            # Carregar o XML
            tree = ET.parse(path)
            root = tree.getroot()

            # Encontrar o elemento 'Aliquota' utilizando o namespace específico
            aliquota_elem = root.find('.//{http://www.betha.com.br/e-nota-contribuinte-ws}Aliquota')
            if aliquota_elem is not None and aliquota_elem.text:
                # Substituir o ponto por vírgula
                return aliquota_elem.text.replace('.', ',')
        except ET.ParseError as e:
            print(f"Erro ao analisar XML: {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar XML: {e}")

        return None

    def get_item_lista_servico_from_xml_nfs(self, num_nota, id_solicitacao):
        path = rf'C:\Users\user\Downloads\xml_{num_nota}_{id_solicitacao}.xml'

        try:
            # Tentativa de encontrar 'ItemListaServico'
            tree = ET.parse(path)
            root = tree.getroot()
            
            item_lista_servico_elem = root.find('.//ItemListaServico')
            if item_lista_servico_elem is not None and item_lista_servico_elem.text:
                return str(int(item_lista_servico_elem.text))[:4]
        except ET.ParseError:
            print("Erro ao analisar 'ItemListaServico'.")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'ItemListaServico': {e}")

        try:
            # Tentativa de encontrar 'cTribNac'
            tree = ET.parse(path)
            root = tree.getroot()
            namespaces = {'ns': 'http://www.sped.fazenda.gov.br/nfse'}

            cTribNac_elem = root.find('.//ns:cTribNac', namespaces)
            if cTribNac_elem is not None and cTribNac_elem.text:
                return cTribNac_elem.text[:4]
        except ET.ParseError:
            print("Erro ao analisar 'cTribNac'.")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'cTribNac': {e}")
        try:
            # Tentativa de encontrar 'codigo_item_lista_servico'
            tree = ET.parse(path)
            root = tree.getroot()
            
            item_lista_servico_elem = root.find('.//codigo_item_lista_servico')
            if item_lista_servico_elem is not None and item_lista_servico_elem.text:
                return str(int(item_lista_servico_elem.text))[:4]
        except ET.ParseError:
            print("Erro ao analisar 'codigo_item_lista_servico'.")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'codigo_item_lista_servico': {e}")
        try:
            # Carregar o XML
            tree = ET.parse(path)
            root = tree.getroot()
            
            # Encontrar o elemento 'ItemListaServico' utilizando o namespace específico
            item_lista_servico_elem = root.find('.//{http://www.betha.com.br/e-nota-contribuinte-ws}ItemListaServico')
            if item_lista_servico_elem is not None and item_lista_servico_elem.text:
                return str(int(item_lista_servico_elem.text))[:4]
        except ET.ParseError as e:
            print(f"Erro ao analisar 'ItemListaServico': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'ItemListaServico': {e}")
        try:
            # Carregar o XML
            tree = ET.parse(path)
            root = tree.getroot()
            
            # Encontrar o elemento 'Codigo'
            codigo_elem = root.find('.//Codigo')
            if codigo_elem is not None and codigo_elem.text:
                # Remover o ponto do código do serviço, se houver
                return codigo_elem.text.replace('.', '')
        except ET.ParseError as e:
            print(f"Erro ao analisar 'Codigo': {e}")
        except Exception as e:
            print(f"Erro desconhecido ao processar 'Codigo': {e}")

        return None
    
    def minimize(self):
        pyautogui.hotkey("win", "m")

    def focus_nbs(self):
        try:
            title = "Barra de Tarefas"
            try:
                app = Application(backend='uia').connect(title=title)
            except ElementNotFoundError:
                print("Não foi possível se conectar com a janela Barra de Tarefas")
                return 
            janela = app[title]
            focar_nbs = janela.child_window(title="SisFin - 1 executando o windows")
            try:
                self.wait_until_interactive(focar_nbs)
            except TimeoutError as e:
                print(str(e))
                return
            focar_nbs.click_input()
        except:
            print("sisfin nao esta aberto")


    def check_if_file_exists_in_downloads(self, num_nota, num_solicitacao, tipodoctonota, download_folder=None):
        if download_folder is None:
            download_folder = str(Path.home() / "Downloads")

        expected_file_name = f"xml_{num_nota}_{num_solicitacao}.xml"
        expected_file_path = os.path.join(download_folder, expected_file_name)

        if not os.path.isfile(expected_file_path):
            return False

        try:
            ET.parse(expected_file_path)
        except ET.ParseError:
            return False
        if tipodoctonota == 'NFS':
            if self.get_verification_code_from_xml_nfs(num_nota, num_solicitacao) is None:
                return False
        elif tipodoctonota == 'NFE':
            return False

        return True


    def get_attempt_count(self, id_solicitacao, numerodocto):
        """Obtém a contagem atual de tentativas para uma dada solicitação e documento."""
        path = r"C:\Users\user\Documents\rpa_project\tentativas"
        filename = f"{id_solicitacao}_{numerodocto}_attempts.txt"
        filepath = os.path.join(path, filename)
        try:
            with open(filepath, "r") as file:
                return int(file.read())
        except FileNotFoundError:
            return 0

    def update_attempt_count(self, id_solicitacao, numerodocto, count):
        """Atualiza a contagem de tentativas no arquivo."""
        path = r"C:\Users\user\Documents\rpa_project\tentativas"
        filename = f"{id_solicitacao}_{numerodocto}_attempts.txt"
        filepath = os.path.join(path, filename)
        with open(filepath, "w") as file:
            file.write(str(count))

    def delete_attempt_file(self, id_solicitacao, numerodocto):
        """Deleta o arquivo de tentativas."""
        path = r"C:\Users\user\Documents\rpa_project\tentativas"
        filename = f"{id_solicitacao}_{numerodocto}_attempts.txt"
        filepath = os.path.join(path, filename)
        os.remove(filepath)
        print(f"Arquivo {filepath} deletado com sucesso.")

    def delete_all_files_in_directory(self, directory):
        for filename in os.listdir(directory):
            file_path = os.path.join(directory, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.remove(file_path)
                    print(f"Arquivo {file_path} deletado com sucesso.")
                elif os.path.isdir(file_path):
                    # Se quiser também deletar diretórios, remova o comentário da próxima linha
                    # shutil.rmtree(file_path)
                    pass
            except Exception as e:
                print(f"Erro ao deletar {file_path}. Causa: {e}")



    def is_within_blocked_time(self):
        current_time = datetime.datetime.now().time()
        blocked_times = [("10:20", "10:35"), ("12:20", "12:35"), ("17:20", "17:35"), ("22:20", "22:35")]

        for start, end in blocked_times:
            start_time = datetime.datetime.strptime(start, "%H:%M").time()
            end_time = datetime.datetime.strptime(end, "%H:%M").time()

            if start_time <= current_time <= end_time:
                return True
        return False
    

    def funcao_main(self):
        registros = database.consultar_dados_cadastro()
        empresa_anterior = None
        for row in registros:
            dir_path = os.path.dirname(os.path.realpath(__file__))
            info_file_path = os.path.join(dir_path, 'info.txt')

            # Removendo o arquivo info.txt se ele existir
            if os.path.exists(info_file_path):
                os.remove(info_file_path)
            self.efetivado = False
            # if not self.is_within_blocked_time():
            id_solicitacao = row[0]
            time.sleep(1)
            database.atualizar_boletos_vencidos(id_solicitacao)
            time.sleep(2)
            database.atualizar_notas_vencidas(id_solicitacao)
            time.sleep(2)
            database.atualizar_adtos_vencidos(id_solicitacao)
            time.sleep(3)
            numerodocto = row[20]
            notas_fiscais = database.consultar_nota_fiscal(id_solicitacao, numerodocto)
            if notas_fiscais:
                # numerodocto = notas_fiscais[0][2]
                serie_value = notas_fiscais[0][3]
                data_emissao_value = notas_fiscais[0][1] 
                tipo_docto_value = notas_fiscais[0][5]
                valor_value = notas_fiscais[0][0]
                inss = notas_fiscais[0][6]
                irff = notas_fiscais[0][7]
                piscofinscsl = notas_fiscais[0][8]
                iss = notas_fiscais[0][9]
                vencimento_value = notas_fiscais[0][4]
            if os.path.exists(info_file_path):
                os.remove(info_file_path)
            time.sleep(2)
            with open('info.txt', 'w') as file:
                file.write(f"{id_solicitacao},{numerodocto}\n")
                file.flush()
                os.fsync(file.fileno())
            # try:
            # tentativas = self.get_attempt_count(id_solicitacao, numerodocto)
            # if tentativas <= 1:
            try:
                # self.update_attempt_count(id_solicitacao, numerodocto, tentativas + 1)
                # self.delete_all_files_in_directory(r"C:\Users\user\Documents\rpa_project\tentativas")
                self.existe_xml = False
                id_empresa = row[18]
                tipo_pagamento_value = row[5]
                numeroos = row[11]
                time.sleep(1)
                self.remover_arquivo_nota(numerodocto, id_solicitacao)
                self.remover_arquivo_xml(numerodocto, id_solicitacao)
                time.sleep(1)
                wise_instance = Wise()
                wise_instance.get_nf_values(numerodocto, id_solicitacao)
                time.sleep(3)
                wise_instance.rename_last_downloaded_file(numerodocto, id_solicitacao)
                time.sleep(2)
                self.xml_existe = self.check_if_file_exists_in_downloads(numerodocto, id_solicitacao, tipo_docto_value)
                print(self.xml_existe)
                # chave_de_acesso_value = self.get_chave_acesso(numerodocto, id_solicitacao)
                time.sleep(1)
                # if tipo_docto_value == 'NFS':
                try:
                    cod_nfse = self.get_codigo_from_pdf(numerodocto, id_solicitacao)
                except:
                    cod_nfse = None
                # serie_nota = self.get_serie(numerodocto, id_solicitacao)
                pdf_inteiro = self.extract_text_from_pdf(numerodocto, id_solicitacao)
                natureza_value = self.find_isolated_5102(pdf_inteiro)
                icms = self.get_valor_icms(numerodocto, id_solicitacao)
                boletos = database.consultar_boleto(id_solicitacao)
                cnpj = row[1]
                time.sleep(1)
                success, message = self.check_conditions(tipo_docto_value, natureza_value, vencimento_value, inss, irff, piscofinscsl, tipo_pagamento_value, icms, boletos, cod_nfse, numeroos, iss)  
                time.sleep(2)
                if success:
                # try:
                    if tipo_docto_value == 'NFE':
                        if self.xml_existe == False:
                            wise_instance.get_xml(cnpj, numerodocto, id_solicitacao)
                            time.sleep(2)
                            chave_de_acesso_value = wise_instance.get_valor_chave_acesso()
                            print(chave_de_acesso_value)
                            time.sleep(1)
                            wise_instance.download_nota()
                            time.sleep(2)
                    elif tipo_docto_value != 'NFE':
                        chave_de_acesso_value = None
                    self.minimize()
                    time.sleep(2)
                    empresa_atual = row[2]
                    self.focus_nbs()
                    time.sleep(2)
                    if empresa_atual != empresa_anterior:
                        self.close_aplications_end()
                        time.sleep(3)
                        self.open_application()
                        self.login()
                        cod_matriz = row[3]
                        self.janela_empresa_filial(row[2], row[3])
                    empresa_anterior = empresa_atual
                    print(empresa_anterior, empresa_atual)
                    self.access_contas_a_pagar()
                    self.janela_entrada(tipo_docto_value)
                    if tipo_docto_value == 'NFE':
                        self.importar_xml()
                        self.abrir_xml(chave_de_acesso_value, numerodocto, id_solicitacao)
                    contab_descricao_value = row[6]
                    cod_contab_value = row[7]
                    total_parcelas_value = row[9]
                    natureza_financeira_value = row[8]
                    usa_rateio_centro_custo = row[14] 
                    valor_sg = row[15] 
                    id_rateiocc = row[16]
                    obs = row[17]
                    cidade_cliente = row[19]
                    rateios_aut = database.consultar_rateio_aut(id_rateiocc)
                    boletos = database.consultar_boleto(id_solicitacao)
                    rateios = database.consultar_rateio(id_solicitacao)
                    num_parcelas = database.numero_parcelas(id_solicitacao, numerodocto)
                    # numeroos = row[11]
                    terceiro = row[12]
                    estado = row[13]
                    time.sleep(3)
                    self.janela_cadastro_nf(cnpj, numerodocto, serie_value, data_emissao_value, tipo_docto_value, valor_value, contab_descricao_value, total_parcelas_value, tipo_pagamento_value, natureza_financeira_value, numeroos, terceiro, estado, usa_rateio_centro_custo, valor_sg, rateios, rateios_aut, inss, irff, piscofinscsl, iss, vencimento_value, obs, cod_contab_value, boletos, num_parcelas, id_solicitacao, chave_de_acesso_value, cod_matriz, cidade_cliente, empresa_atual)
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
                    wise_instance.Anexar_AP(id_solicitacao, num_controle, numerodocto, serie_value)
                    time.sleep(2)
                    wise_instance.get_pdf_file(numerodocto, id_solicitacao)
                    time.sleep(4)
                    wise_instance.confirm()
                    time.sleep(3)
                    # send_mail.enviar_email(id_solicitacao, numerodocto)
                    time.sleep(4)
                    self.back_to_nbs()
                    time.sleep(2)
                    self.close_aplications_half()
                    time.sleep(1)
                    # self.delete_all_files_in_directory(r"C:\Users\user\Documents\rpa_project\tentativas")
                    time.sleep(2)
                    # registro_existe = database.verificar_dados_cadastro(numerodocto, id_solicitacao)
                    time.sleep(2)
                    ap_existe = rf"C:\Users\user\Documents\APs\AP_{numerodocto}{id_solicitacao}.pdf"
                    if os.path.exists(ap_existe):
                        database.atualizar_anexosolicitacaogasto(numerodocto, id_solicitacao)
                        time.sleep(3)
                        if not self.check_cod_nfs(tipo_docto_value, cod_nfse) and serie_value not in ('TL', 'CO', 'FA', 'CCF', None):
                            self.send_success_message_cod_nfs_exception(id_solicitacao, numerodocto, tipo_docto_value)
                        elif not self.check_tribut(inss, irff, piscofinscsl, iss) and serie_value not in ('TL', 'CO', 'FA', 'CCF', None):
                            self.send_success_message_cod_tribut_exceptions(id_solicitacao, numerodocto, tipo_docto_value)
                        elif not self.check_tribut(inss, irff, piscofinscsl, iss) and not self.check_cod_nfs(tipo_docto_value, cod_nfse) and serie_value not in ('TL', 'CO', 'FA', 'CCF', None):
                            self.send_success_message_cod_tribut_cod_nfs_exceptions(id_solicitacao, numerodocto, tipo_docto_value)
                        elif serie_value == 'TL':
                            self.send_success_message_telecomunication(id_solicitacao, numerodocto, tipo_docto_value)
                        elif serie_value in ('CO', 'CCF'):
                            self.send_success_message_compromisso(id_solicitacao, numerodocto, tipo_docto_value)
                        elif serie_value == 'FA':
                            self.send_success_message_fatura(id_solicitacao, numerodocto, tipo_docto_value)
                        elif serie_value is None:
                            self.send_success_message_not_serie(id_solicitacao, numerodocto, tipo_docto_value)
                        else:
                            self.send_success_message(id_solicitacao, numerodocto, tipo_docto_value)
                    else:
                        self.send_ap_nao_existe_message(id_solicitacao, numerodocto, tipo_docto_value)
                        # continue
                    # except:
                    #     self.send_message_with_traceback(id_solicitacao, numerodocto)
                else:
                    self.send_message_pre_verification(message, id_solicitacao, numerodocto)
                    database.autoriza_rpa_para_r(id_solicitacao)
                    time.sleep(2)
                    self.back_to_nbs()
            except Exception as e:
                print(traceback.format_exc()) 
                time.sleep(1)
                database.autoriza_rpa_para_r(id_solicitacao)
                time.sleep(2)
                self.send_failure_try_message(id_solicitacao, numerodocto)
                raise
                # self.delete_all_files_in_directory(r"C:\Users\user\Documents\rpa_project\tentativas")




            
rpa = NbsRpa()
rpa.funcao_main()
