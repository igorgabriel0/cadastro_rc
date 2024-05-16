from math import isnan
import math
import win32com.client
import sys
import subprocess
import time
import datetime
import pandas as pd
import pyautogui as auto
import pyperclip
import os
import shutil as stl
from openpyxl import workbook, load_workbook

contador = 0
session = ""

usuario_logado = os.getlogin()
#user_vm = f"C:\Users\{usuario_logado}\JSL SA\Grupo Vamos - Indicadores - Automação RC - Automação RC\automação"
#print(usuario_logado)

# VARRE A PASTA DE ARQUIVOS
def get_files(dir=None,return_dir = False):
    list = []
    for file in os.listdir(dir):
        if file.endswith(".xlsx"):
            file_path = f"{file}"
            if return_dir == True:
                list.append(str(dir+"/"+file_path))
            else:
                list.append(str(file_path))
    return (list)

# VARRE A PASTA DE ACESSOS
def get_login(dir=None,return_dir = False):
    list = []

    for file in os.listdir(dir):
        if file.endswith(".xlsx"):
            file_path = f"{file}"
            if return_dir == True:
                list.append(str(dir+"/"+file_path))
            else:
                list.append(str(file_path))
    return (list)

lista_login = get_login(dir=r"C:\Users\igor.gabriel\OneDrive - JSL SA\Automação RC\ACESSOS", return_dir= True) # Pasta de Logins
lista_arquivos = get_files(dir=r"C:\Users\igor.gabriel\OneDrive - JSL SA\Automação RC\REQUISIÇÕES", return_dir= True) # Pasta Modelos de RC

# Le os arquivos dentro da pasta REQUISIÇÕES
for file in lista_arquivos:
    try:
        print(f"FILE: {file}")
        dt = pd.read_excel(file, sheet_name= 'Gerar RC')
        for index, linha in dt.iterrows():
            user = dt.at[0, 'Usuário'] # USUÁRIO
            pasta_acessos = r'C:\Users\igor.gabriel\OneDrive - JSL SA\Automação RC\ACESSOS' # Caminho para a pasta 'Acessos'
            arquivos_acessos = os.listdir(pasta_acessos) # Lista todos os arquivos na pasta 'Acessos'

        # Procura por uma planilha com o nome correspondente ao valor da variável user
        planilha_encontrada = None
        for arquivo in arquivos_acessos:
            if arquivo.startswith(user) and arquivo.endswith('.xlsx'):
                planilha_encontrada = os.path.join(pasta_acessos, arquivo)
                la = pd.read_excel(planilha_encontrada, sheet_name = 'Planilha1')
                login = la.at[0, 'login']
                senha = la.at[0, 'senha']
                print(user)
                print(f"{login} {senha}")

        def saplogin(): # Função de Login
            global session
            try:
                path = r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe" # Define o caminho para o SAP
                subprocess.Popen(path) # Inicia o SAP
                time.sleep(3) # Espera 5 segundos para que o SAP GUI seja carregado
                SapGuiAuto = win32com.client.GetObject('SAPGUI') # Obtém a a\utomação do SAP GUI
                if not type(SapGuiAuto) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado
                    return
                application = SapGuiAuto.GetScriptingEngine # Obtém a aplicação SAP GUI
                if not type(application) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado        
                    SapGuiAuto = None
                    return

                if user == "igor.gabriel":
                    print("CONECTION: Qualidade")
                    connection = application.OpenConnection("$ JSL - ECC - Qualidade (ECQ)", True) # Conexão Qualidade(TESTES)
                else:
                    print("CONECTION: Produção")
                    connection = application.OpenConnection("# JSL -  ECC - Produção (ECP)", True) # Conexão Produção
                
                #connection = application.OpenConnection("$ JSL - ECC - Qualidade (ECQ)", True)

                if not type(connection) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado
                    application = None
                    SapGuiAuto = None
                    return
                session = connection.Children(0) # Obtém a sessão ativa
                if not type(session) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado
                    connection = None
                    application = None
                    SapGuiAuto = None
                    return
                # Preenche os campos de nome de usuário e senha no SAP GUI
                session.findById("wnd[0]/usr/txtRSYST-BNAME").text = f"{login}" # Preenche a matricula no campo de login
                session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = f"{senha}" # Preenche a senha
                #auto.hotkey('enter') # Pressiona enter para completar o login
            except:
                print(sys.exc_info()[0]) # Em caso de exceção, aparece mensagem de erro
            finally:
                # Limpa os objetos e recursos   
                connection = None
                application = None
                SapGuiAuto = None
        # Chama a função saplogin para executar o login no SAP GUI
        saplogin()
        print("ABRINDO O SAP...")

        def get_connection(ambiente = "# JSL -  ECC - Produção (ECP)"): # Função que define a conexão no SAP
            SapGuiAuto = win32com.cliente.GetObject('SAPGUI') 
            application = SapGuiAuto.GetScriptingEngine 
            connection = application.OpenConnection(ambiente, True)
            return connection

        def get_session(connection):
            session = connection.Children(0)
            return session

        def pesquisa():
            session.findById("wnd[0]/tbar[0]/okcd").text = "me51n"
            auto.hotkey('enter')
        pesquisa()

        def set_tipoRC(tipo_rc = '', row = 0):
            auto.hotkey('Ctrl','F2')
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:3327/cmbMEREQ_TOPLINE-BSART").key = tipo_rc
                    return
                except:
                    pass

        def set_codigo(material = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","MATNR",f"{material}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "MATNR"
                    return
                except:
                    pass

        def set_cabecalho(texto = '', row = 0):
            if texto != None or texto != "":
                for i in range(10,20,1):
                    try:
                        session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text = f"{texto}"
                        time.sleep(2)
                    except:
                        pass     

        def set_centro(centro = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","NAME1", f"{centro}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "NAME1"
                    auto.hotkey('enter')
                    return
                except:
                    pass

               

        def set_qtd(qtd = 0, row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","MENGE",f"{qtd}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "MENGE"
                    auto.hotkey('enter')
                    return
                except:
                    pass

        def set_texto(texto = '', row = ''):
            if texto != None or texto != "":
                txt = auto.locateOnScreen('TEXTO.png', confidence=0.9)
                if txt is not None:
                    time.sleep(3)
                    auto.click(auto.locateCenterOnScreen('TEXTO.png', confidence=0.9))
                    print("\nTEXTO\n")
                else:
                    print("TEXTO NAO TA NA TELA")
                    aba_item = auto.locateCenterOnScreen('DETALHE.png', confidence=0.9)
                    if aba_item is not None:
                        print("ABRE ABA DE ITEM")
                        time.sleep(2)
                        auto.click(auto.locateCenterOnScreen('DETALHE.png', confidence=0.9))
                        time.sleep(1)
                        auto.click(auto.locateCenterOnScreen('DETALHE.png', confidence=0.9))
                        time.sleep(7)
                        auto.click(auto.locateCenterOnScreen('TEXTO.png', confidence=0.9))
                    else:
                        pass
                print("ABRE CAIXA TEXTO")
                time.sleep(2)
                auto.click(auto.locateCenterOnScreen('CAIXA.png', confidence=0.9))
                time.sleep(2)
                for i in range(10,20,1):
                    try:                       
                        session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").setSelectionIndexes (5,5)
                        session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT13/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1329/subTEXTS:SAPLMMTE:0200/subEDITOR:SAPLMMTE:0201/cntlTEXT_EDITOR_0201/shellcont/shell").text = f"{texto}"
                    except:
                        pass
            else:
                pass

        def set_preco_av(preco = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","PREIS", f"{preco}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "PREIS"
                    auto.hotkey('enter')
                    return
                except:
                    pass
 
        def set_deposito(deposito = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","LGOBE", f"{deposito}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "LGOBE"
                    auto.hotkey('enter')
                    return
                except:
                    pass

        def set_class(class_cont = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","KNTTP", f"{class_cont}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "KNTTP"
                    time.sleep(2)
                    auto.hotkey('enter')
                except:
                    pass

        def set_gp_comp(grp_comp = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","EKGRP",f"{grp_comp}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").setCurrentCell (0,"EKGRP")
                    auto.hotkey('enter')
                except:
                    pass

        def set_fornecedor(forn = '', row =0):
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","FLIEF",f"{forn}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "FLIEF"
                    auto.hotkey('enter')
                except:
                    pass

        def set_cc(cc = '', row = 0):
            time.sleep(2)
            auto.click(auto.locateCenterOnScreen('BAIXO.png', confidence=0.9))
            time.sleep(2)
            auto.click(auto.locateCenterOnScreen('DADOS.png', confidence=0.9))
            time.sleep(2)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZCENTROCUSTO").text = f"{cc}"
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZCENTROCUSTO").setFocus
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZCENTROCUSTO").caretPosition = 10
                except:
                    pass

        def set_cc_cic(cc = '', row = 0):
            time.sleep(2)
            # ABRE ABA CC
            auto.click(auto.locateCenterOnScreen('BAIXO.png', confidence=0.9))
            time.sleep(2)
            centroc = auto.locateCenterOnScreen('CC.png', confidence=0.9)
            if centroc is not None:
                auto.click(auto.locateCenterOnScreen('CC.png', confidence=0.9))
            else:
                pass
            time.sleep(5)
            # MUDA PARA MODO CC SIMPLES
            simples = auto.locateCenterOnScreen('SIMPLES.png', confidence=0.9)
            simples_s = auto.locateCenterOnScreen('SIMPLES_S.png', confidence=0.9)
            if simples is not None:
                auto.click(auto.locateCenterOnScreen('SIMPLES.png',confidence= 0.9))
            elif simples_s is not None:
                auto.click(auto.locateCenterOnScreen('SIMPLES_S.png',confidence= 0.9))
            else:
                pass
            print("ABRE CC SIMPLES")
            time.sleep(3)

            for i in range(10, 20, 1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").text = f"{cc}"
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1100/subKONTBLOCK:SAPLKACB:1101/ctxtCOBL-KOSTL").caretPosition = 10
                    auto.hotkey('enter')
                except:
                    pass

        def set_nf(data_nf = '', data_ven = '', n_nf = '', row = 0):
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15").select
                    print("ABRE A ABA DE NF")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZZDATA_NF").text = f"{data_nf}"
                    auto.hotkey('tab')
                    print("PREENCHE A DATA")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZZDATA_VCTO_NF").text = f"{data_ven}"
                    auto.hotkey('tab')
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/txtCI_EBANDB-ZZNRO_NF").text = f"{n_nf}"
                    auto.hotkey('enter')                
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/txtCI_EBANDB-ZZNRO_NF").setFocus
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/txtCI_EBANDB-ZZNRO_NF").caretPosition = 3
                    auto.hotkey('enter')
                except:
                    pass

        def set_contrato(contrato = '', row = 0):
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","KONNR",f"{contrato}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "KONNR"
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").firstVisibleColumn = "LGOBE"
                    auto.hotkey("enter")
                except:
                    pass

        def set_item(item = '', row = 0):
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","KTPNR",f"{item}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "KTPNR"
                    auto.hotkey('enter')
                except:
                    pass
        
        def abaerro():
            print("ABA ERRO")
            time.sleep(2)
            auto.hotkey('Shift','F4')
            time.sleep(1) 
            auto.hotkey('Shift','F6')
            time.sleep(1)
            auto.hotkey('Shift','F7')
            time.sleep(1)
            auto.hotkey('tab')
            time.sleep(1)
            auto.hotkey('space')
            auto.hotkey('Ctrl','Shift','F7')
            time.sleep(2)

        def retorno():
            global msg_erro
            print("\nABRE MENSAGEM RETORNO...\n")
            time.sleep(2)
            msg_rc = session.findById("wnd[0]/sbar").Text # Salva informação da barra inferior
            requisicao = "Requisição de compra criada sob nº" 
            if requisicao in msg_rc and msg_erro != "":
                print("LIMPOU VARIAVEL ERRO")
                msg_erro = ""
            else:
                pass
            print(f"MENSAGEM RC: {msg_rc}")
            print(f"MSG_ERRO: {msg_erro}")
            if msg_erro != "":
                # --- PASTA ERROS ---
                print(f"POSSUI ERRO | Mensagem ERRO: {msg_erro}")
                user = dt.iloc[0]["Usuário"] # DEFININDO USUÁRIO
                erro = "ERROS"
                if not os.path.exists(erro): # Cria pasta ERROS se não existir
                    os.makedirs(erro)
                    print(f"Pasta {erro} criada com sucesso!")

                erro_user_path = f"{erro}/ERROS_{user}" # ERRO POR USUÁRIO 
                print(erro_user_path)
                if not os.path.exists(erro_user_path): # VERIFICAÇÃO E CRIAÇÃO DE PASTA DE ERRO POR USUÁRIO
                    os.makedirs(erro_user_path)
                    print(f"Pasta {erro_user_path} criada com sucesso")

                data = {'Chave': key, 'Mensagem Retorno': msg_erro}
                df_aux = pd.DataFrame(data, index=[index])  # Cria excel com mensagem de retorno e adiciona quantas linhas desejar
                df_erro = pd.concat([df_aux])

                # Limpando o nome do arquivo
                nome_arquivo = file.split('/')[-1]
                nome_arquivo = nome_arquivo.split('.')[0]

                df_erro.to_excel(f'{erro_user_path}/ERRO_{nome_arquivo}_{key}_{hora_inicio}_.xlsx') # Definindo local e nome do arquivo
                print(f"CRIANDO ARQUIVO ERRO: {user}_{hora_inicio}_{nome_arquivo}_{key}_.xlsx")
                df_erro = "" # Limpando variável
            
            # --- CRIA REQUISIÇÃO COM ARQUIVO ---
            else:
                if msg_rc == "":
                    msg_rc = "Erro não identificado"
                else:
                    pass
                print(msg_rc)
                data = {'Chave': key, 'Mensagem Retorno': msg_rc}
                df_aux = pd.DataFrame(data, index=[index, index + 1])
                df_log = pd.concat([df_aux])
                print(df_log)

                # --- PASTA LOG ---
                user = dt.iloc[0]["Usuário"] # DEFININDO USUÁRIO  
                pasta = "LOG" # NOMEANDO PASTA COMO LOG
                
                if not os.path.exists(pasta): # Cria pasta LOG se não existir
                    os.makedirs(pasta)
                    print(f"Pasta '{pasta}' criada com sucesso!")
                print("PASTA LOG EXISTE")
                log_user = f"LOG/LOG_{user}" # LOG POR USUÁRIO
                
                if not os.path.exists(log_user): # Cria pasta user log se não existir
                    os.makedirs(log_user)
                    print(f"Pasta {log_user} criada com sucesso!")

                nome_arquivo = file.split('/')[-1] # Limpando o nome do arquivo
                nome_arquivo = nome_arquivo.split('.')[0]

                df_log.to_excel(f"{log_user}/LOG_{hora_inicio}_{nome_arquivo}_{key}_.xlsx") # Definindo local e nome do arquivo
                print(file)

                # === REMOVE ARQUIVO ===
                #if msg_rc != "Erro não identificado":
                #    os.remove(file)
                #    print(f"\nREMOVE ARQUIVO {file}\n")
                #else:
                #    pass
                #    
                df_log = "" # Limpa variável df_log 
                print(f"{log_user}/LOG_{hora_inicio}_{nome_arquivo}.xlsx")
                print(f"Mensagem retornada: {msg_rc} | CHAVE: {key} ")

        # Pegando informação de data e hora
        hora_inicio = datetime.datetime.now()
        hora_inicio = hora_inicio.strftime("%d-%m-%Y_%H-%M")

        # ---DIRETORIO PANDAS ---
        df_log = pd.DataFrame()
        df_erro = pd.DataFrame()
        print(f"TODOS AS PLANILHAS: {lista_arquivos}")
        dt = pd.read_excel(file, sheet_name= 'Gerar RC')

        # --- DEFININDO USUÁRIO ---
        user = dt.iloc[0]["Usuário"]

        #---CONTADORES---#
        coluna = dt['Chave'] # Coluna Chave da planilha
        cont = len(coluna) # Contador de inserção 
        index = 0 # Contador de linha(Planilha)
        row = 0 #Contador de linha(SAP)

        # ╰(*°▽°*)╯ LOOPS DE INSERÇÃO (^///^)
        for index, linha in dt.iterrows(): # Le linha a linha da planilha  
            dt['Chave'] = dt['Chave'].where(pd.notnull(dt['Chave']), None) # CONVERTE A COLUNA CHAVE EM NONE O QUE FOR NAN
            print(f"CONVERTEU PARA NONE  PLANILHA: {file}")
            print("INICIA O FOR DA LINHA")
            chave = linha['Chave']
            print(f"CHAVE: {chave}")
            if pd.isna(chave):
                chave = None
                print(f"Converte CHAVE para None: {chave}")
                pass
            else:
                key = ''.join(chave.split("-") and chave.split("/") and chave.split(" ")).replace(" ", "")
                key = str(key)

            print(f"KEY : {key}")
            tipo_rc = linha['Tipo RC']
            material = str(linha['Código'])

            if "-" in material:
                material = str(material)
                print(f"CODIGO {material} STRING")
            else:
                material = int(float(material))
                print(f"CODIGO {material} INTEIRO")

            texto = linha['Texto na RC']
            centro = linha['Centro']
            qtd = linha['Qtd']
            preco_av = linha['Preço Avaliação']
            deposito = int(linha['Depósito'])
            class_cont = linha['Class Cont']
            grp_comp = linha['Grp Comp']
            forn = linha['Cód Forn']
            cc = linha['C/C']
            data_nf = linha['Data NF']
            data_ven = linha['Data ven']
            n_nf = linha['N NF']
            contrato = linha['Contrato']
            item = linha['It Contrato']
            anexo = linha['Anexo']
            
            print(f"KEY: {key} {row}| TIPO RC: {tipo_rc} {row}| CODIGO: {material} {row}| CENTRO: {centro} {row}| QUANTIDADE: {qtd} {row}| PREÇO: {preco_av} {row}| DEPOSITO: {deposito} {row}| CLASS: {class_cont} {row}| CENTRO DE CUSTO: {cc} {row}")
            
            # IDENTIFICAR PROXIMA LINHA
            if index < len(dt) - 1:
                proxima_key = dt.loc[index + 1, 'Chave']
            else:
                proxima_key = None

            # (☞°ヮ°)☞ CHAMAR FUNÇÕES ☜(°ヮ°☜)
            set_tipoRC(tipo_rc)
            set_codigo(material, row = contador)
            set_cabecalho(texto, row = contador)
            set_centro(centro, row = contador)
            set_qtd(qtd, row = contador)
            set_texto(texto)
            set_class(class_cont, row = contador)
            set_gp_comp(grp_comp, row = contador)
            set_deposito(deposito, row = contador)
            
            # --- CENTRO DE CUSTO ---
            if class_cont == 'K' or class_cont == 'k': # CENTRO DE CUSTO CIC
                print(f"\n RC {tipo_rc} CIC \n")
                set_cc_cic(cc, row = contador)
            else:
                print(f"\n RC {tipo_rc} \n") # CENTRO DE CUSTO NORMAL
                set_cc(cc, row = contador)
            
            if pd.isna(preco_av):
                print("ISNAN PREÇO")
                pass
            else:
                set_preco_av(preco_av, row = contador)
            
            # SEGUNDA TENTATIVA CENTRO DE CUSTO
            msg_erro = session.findById("wnd[0]/sbar").Text
            print(f"\nMSG_ERRO: {msg_erro}\n")
            if msg_erro == "Entrar Centro custo":
                set_cc_cic(cc, row = contador)
            else:
                pass
            # --- CÓDIGO DE FORNECEDOR ---
            if pd.notnull(forn):
                set_fornecedor(forn, row = contador)
            else:
                pass

            # --- NF ---
            if tipo_rc == "ZRE":
                set_nf(data_nf, data_ven, n_nf, row = contador)
            else:
                msg_erro == ""

            # --- CONTRATO | ITEM ---
            if pd.notnull(contrato and item):
                print("CONTRATO | ITEM")
                set_contrato(contrato, row = contador)
                set_item(item, row = contador)
            else:
                pass

            msg_erro1 = session.findById("wnd[0]/sbar").Text
            print(f"\nMSG_ERRO 2: {msg_erro1}\n")
            if msg_erro1 != "" and msg_erro == "":
                msg_erro = msg_erro1
            else:
                pass

            # --- ANEXO ---
            if pd.notnull(anexo): # Verifica se a célula da coluna arquivo não está vazia
                pasta_anexos = r'C:\Users\igor.gabriel\OneDrive - JSL SA\Automação RC\anexos'
                arquivos_anexo = os.listdir(pasta_anexos)
                for anexo_user in arquivos_anexo:
                    if anexo_user == user:
                        anexos = os.path.join(pasta_anexos, anexo_user)
                        pasta_anexo_user = os.listdir(anexos)
                        print(f"PASTA_ANEXO_USER: {pasta_anexo_user}")
                if anexo in pasta_anexo_user:
                    print(f"anexo: {anexo} esta em {pasta_anexo_user}")

                    anexo_str = str(anexo)
                    divisao = anexo_str.split(';')
                    conta = 0
                    while conta < len(divisao):
                        try:
                            time.sleep(3)
                            session.findById("wnd[0]/titl/shellcont/shell").pressContextButton ("%GOS_TOOLBOX")
                            session.findById("wnd[0]/titl/shellcont/shell").selectContextMenuItem ("%GOS_PCATTA_CREA")
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = anexos
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = f"{divisao[conta]}"
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
                            auto.hotkey('enter')
                            time.sleep(2)
                            i = 0

                            while i < 2:
                                auto.press('tab', presses= 4)
                                auto.press('enter', presses= 2)
                                time.sleep(2)
                                i += 1
                            conta += 1 
                        except:
                            pass
                else:
                    print(f"anexo: {anexo} NÃO esta em {pasta_anexo_user}")
                    pass
            else:
                pass

            time.sleep(10)
            contador += 1
            print(f"PROXIMA KEY: {proxima_key}\n")
            print(f"KEY: {key} {row}| TIPO RC: {tipo_rc} {row}| CODIGO: {material} {row}| CENTRO: {centro} {row}| QUANTIDADE: {qtd} {row}| PREÇO: {preco_av} {row}| DEPOSITO: {deposito} {row}| CLASS: {class_cont} {row}| CENTRO DE CUSTO: {cc} {row}")
            
            # === PROCESSO DE FINALIZAÇÃO ===
            if proxima_key != key and proxima_key is not None: # Verifica se a proxima key é diferente e se não é None
                print("---PROXIMA LINHA---")
                time.sleep(2)
                auto.hotkey('Ctrl','s') #SALVA
                print("SALVO")
                time.sleep(2)
                posicao = auto.locateOnScreen('MENSAGEM.png', confidence=0.9) # Localize a posição da imagem de erro
                if posicao is not None: # Verifique se a imagem de erro foi encontrada
                    time.sleep(15)
                    print("\nA imagem está na tela na posição:\n", posicao)
                    if msg_erro == "": # Se não ocorreu algum erro
                        time.sleep(3)
                        auto.hotkey('F12') # Fecha a tela de errp
                        time.sleep(2)
                        auto.click(auto.locateCenterOnScreen('ERRO.png', confidence=0.9)) # Clica no botão erro
                        abaerro() # Abre a função que pega o valor do erro
                        auto.click(x=290, y=179) # Clica dentro do campo de texto
                        auto.hotkey('Ctrl','a') # Seleciona tudo
                        auto.hotkey('ctrl', 'c') # Copia
                        time.sleep(1)
                        msg_erro = pyperclip.paste() 
                        linhas = msg_erro.splitlines()[2:3] 
                        msg_erro = '\n'.join(linhas) # Junta as linhas de volta em um único texto
                        retorno()
                    else:
                        #retorno()
                        pass   
                else:
                    print("\nA imagem não está na tela.\n")
                    pass

                # --- VERIFICAÇÃO SALVAMENTO COM ANEXO ---
                if anexo != "" and msg_erro == "":
                    print(f"CHAVE: {key} (SALVO COM ANEXO)\n")  
                    time.sleep(3)
                    permitir = auto.locateCenterOnScreen('PERMITIR.png', confidence=0.9)
                    if permitir is not None:
                        auto.click(auto.locateCenterOnScreen('PERMITIR.png', confidence=0.9)) # Clica no botão de permitir
                    else:
                        pass
                    time.sleep(2)
                    if permitir is not None:
                        auto.click(auto.locateCenterOnScreen('PERMITIR.png', confidence=0.9)) # Clica no botão de permitir
                    else:
                        pass
                    time.sleep(2)
                    if permitir is not None:
                        auto.click(auto.locateCenterOnScreen('PERMITIR.png', confidence=0.9)) # Clica no botão de permitir
                    else:
                        pass
                    time.sleep(5)
                    retorno()
                    contador = 0
                else:
                    auto.hotkey('F6')
                    time.sleep(3)
                    auto.hotkey('tab')
                    auto.hotkey('enter')
                    print("ABRIU NOVA REQUISIÇÃO")
                    contador = 0

            elif proxima_key is None:
                print("--- FECHANDO RC ---")
                time.sleep(3)
                auto.hotkey('Ctrl','s')
                time.sleep(3)
                contador = 0
                print("SALVOU")

                # Localize a posição da imagem na tela
                posicao = auto.locateOnScreen('MENSAGEM.png', confidence=0.9)
                # Verifique se a imagem foi encontrada
                if posicao is not None:
                    print("\nA imagem está na tela na posição:\n")
                    time.sleep(3)
                    auto.hotkey('F12')
                    time.sleep(2)
                    auto.click(auto.locateCenterOnScreen('ERRO.png', confidence=0.9))
                    abaerro()
                    auto.click(x=290, y=179)
                    auto.hotkey('Ctrl','a')
                    auto.hotkey('ctrl', 'c')
                    time.sleep(1)
                    msg_erro = pyperclip.paste() 
                    linhas = msg_erro.splitlines()[2:3] # Divide o texto em linhas e seleciona as três primeiras linhas
                    msg_erro = '\n'.join(linhas) # Junta as linhas de volta em um único texto
                    retorno()
                    
                else:
                    print("\nA imagem não está na tela.\n")
                    retorno()

            #elif (isinstance(proxima_key, float) and math.isnan(proxima_key)):
            elif pd.isna(proxima_key):
                print("IS NAN")
                time.sleep(3)
                auto.hotkey('Ctrl','s')
                time.sleep(3)
                contador = 0
                print("SALVOU")
                retorno()
            
            elif proxima_key != key :
                print("--- PROXIMA RC ---")
                time.sleep(3)
                auto.hotkey('Ctrl','s')
                time.sleep(3)
                contador = 0
                auto.hotkey('F12')
                time.sleep(1)
                auto.hotkey('F6')
                time.sleep(1)
                auto.hotkey('tab')
                auto.hotkey('enter')
                retorno()

                # --- VERIFICAÇÃO SALVAMENTO COM ANEXO ---
                if anexo != "" and msg_erro == "": 
                    time.sleep(1)
                    print(f"CHAVE: {key} (SALVO COM ANEXO)\n")
                    permitir = auto.locateCenterOnScreen('PERMITIR.png', confidence=0.7)
                    if permitir is not None:
                        auto.click(auto.locateCenterOnScreen('PERMITIR.png', confidence=0.7))
                        time.sleep(2)
                    else:
                        pass
                    if permitir is not None:
                        auto.click(auto.locateCenterOnScreen('PERMITIR.png', confidence=0.7))
                        time.sleep(2)
                    else:
                        pass
                    if permitir is not None:
                        auto.click(auto.locateCenterOnScree('PERMITIR.png', confidence=0.7))
                        time.sleep(2)
                    else:
                        pass
                    retorno()
                else:
                    pass
            else:
                pass
        # --- FECHA O SAP --- 
        print("FINALIZOU REQUISIÇÃO")
        print("FECHANDO SAP...")
        time.sleep(3)
        auto.hotkey('alt','F4')
        time.sleep(2)
        auto.hotkey('alt','F4')
        time.sleep(2)
        auto.hotkey('tab')
        auto.hotkey('enter')
        time.sleep(2)
        auto.hotkey('alt','F4')
        time.sleep(3)
        print("FECHADO")

    # SE OCORRER ERRO AO LER PLANILHA
    except Exception as e:
        print(f"\nErro ao ler o arquivo (ULTIMO EXCEPT): {file}: {e}")  
        print("FECHANDO SAP...")
        time.sleep(3)
        auto.hotkey('alt','F4')
        time.sleep(2)
        auto.hotkey('alt','F4')
        time.sleep(2)
        auto.hotkey('tab')
        auto.hotkey('enter')
        time.sleep(2)
        auto.hotkey('alt','F4')
        time.sleep(3)    
print("... FIM ...")




# Descrição do Projeto
#Este projeto tem como objetivo, automatizar o processo de geração de requisição das equipes, frete e fornecedores.
#Utilizando pastas no canal do Teams “Automação RC”, a automação le todas as planilhas e as executa em sequencia por usuário.
#Dentro do canal existem 4 pastas principais:
#REQUISIÇÕES: onde as planilhas são enviadas e a automação faz a leitura
#ACESSOS: onde é armazenado as informações de login dos usuários que a automação usa para acessar (a pasta é protegida e só a automação le)
#ANEXOS: é armazenado os anexos que são incluidos nas requisições
#LOG: retorna o numero das requisições para os usuários
#ERRO: retorna caso haja erro e qual o erro  




