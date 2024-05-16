from math import isnan
import math
import win32com.client
import sys
import subprocess
import time
import datetime
import pandas as pd
import pyautogui as auto
import os
import shutil as stl
from openpyxl import workbook, load_workbook


# === EQUIPE TEAM ===

contador = 0
session = ""

usuario_logado = os.getlogin()
#user_vm = f"C:\Users\{usuario_logado}\JSL SA\Grupo Vamos - Indicadores - Automação RC - Automação RC\automação"
print(usuario_logado)



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
lista_arquivos = get_files(dir=r"REQUISIÇÕES", return_dir= True) # Pasta Modelos de RC

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
                time.sleep(2) # Espera 5 segundos para que o SAP GUI seja carregado
                SapGuiAuto = win32com.client.GetObject('SAPGUI') # Obtém a a\utomação do SAP GUI
                if not type(SapGuiAuto) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado
                    return
                application = SapGuiAuto.GetScriptingEngine # Obtém a aplicação SAP GUI
                if not type(application) == win32com.client.CDispatch: # Verifica se o objeto obtido é do tipo esperado        
                    SapGuiAuto = None
                    return
                #

                if user == "igor.gabriel":
                    print("CONECTION: Qualidade")
                    connection = application.OpenConnection("$ JSL - ECC - Qualidade (ECQ)", True) # Conexão Qualidade(TESTES)
                else:
                    print("CONECTION: Produção")
                    connection = application.OpenConnection("# JSL -  ECC - Produção (ECP)", True) # Conexão Produção
                
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
                auto.hotkey('enter') # Pressiona enter para completar o login
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

        def set_tipoRC(tipo_rc = '', row = ''):
            #session.findById("wnd[0]").sendVKey (26) # Nota de cabeçalho
            auto.hotkey('cntrl','F2')
            time.sleep(1)
            # COM CABEÇALHO = SAPLMEGUI:0013
            # SEM CABEÇALHO = SAPLMEGUI:0016 
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

        def set_centro(centro = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","NAME1", f"{centro}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "NAME1"
                    #session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
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
            for i in range(10,20,1):
                try:
                    # NOTA DE CABEÇALHO
                    
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3102/tabsREQ_HEADER_DETAIL/tabpTABREQHDT1/ssubTABSTRIPCONTROL3SUB:SAPLMEGUI:1230/subTEXTS:SAPLMMTE:0100/subEDITOR:SAPLMMTE:0101/cntlTEXT_EDITOR_0101/shellcont/shell").text = "teste"
                    time.sleep(3)

                    print("\nTEXTO\n")
                    # Selecionando a aba Textos
                    print("REDIMENCIONANDO")
                    session.findById("wnd[0]").resizeWorkingPane(218,35,False)
                    time.sleep(2)
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/ctxtMEPO1210-EKORG").caretPosition = 2
                    time.sleep(2)
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL").selectTab("TABREQDT13")
                    time.sleep(2)
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/tabsHEADER_TABSTRIP1/tabpTABS_O-0100").select()



                except:
                    pass

        def set_preco_av(preco = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","PREIS", f"{preco}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "PREIS"
                    #session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
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
                    #session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
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
                    #session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0010/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").pressEnter
                    auto.hotkey('enter')
                except:
                    pass

        def set_fornecedor(forn = '', row =0):
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").modifyCell (f"{row}","FLIEF",f"{forn}")
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "FLIEF"
                    #session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").firstVisibleColumn = "EEIND"
                    auto.hotkey('enter')
                except:
                    pass

        def set_cc(cc = '', row = 0):
            time.sleep(2)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15").select
                    time.sleep(2)
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZCENTROCUSTO").text = f"{cc}"
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZCENTROCUSTO").setFocus
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT15/ssubTABSTRIPCONTROL1SUB:SAPLMEGUI:1318/ssubCUSTOMER_DATA_ITEM:SAPLXM02:0111/tabsABAS/tabpABAS_F1/ssubSUB1:SAPLXM02:0112/ctxtCI_EBANDB-ZCENTROCUSTO").caretPosition = 10
                    auto.hotkey('enter')
                except:
                    pass

        def set_cc_cic(cc = '', row = 0):
            time.sleep(1)
            for i in range(10,20,1):
                try:
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:3212/cntlGRIDCONTROL/shellcont/shell").currentCellColumn = "ZCENTROCUSTO"
                    session.findById(f"wnd[0]/usr/subSUB0:SAPLMEGUI:00{i}/subSUB3:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1301/subSUB2:SAPLMEGUI:3303/tabsREQ_ITEM_DETAIL/tabpTABREQDT6/ssubTABSTRIPCONTROL1SUB:SAPLMEVIEWS:1101/subSUB2:SAPLMEACCTVI:0100/subSUB1:SAPLMEACCTVI:1000/tblSAPLMEACCTVIDYN_1000TC/ctxtMEACCT1000-KOSTL[5,0]").text = f"{cc}"
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

        
        def retorno():
            time.sleep(2)
            msg_rc = session.findById("wnd[0]/sbar").Text # Salva informação da barra inferior
            print(f"MENSAGEM RC: {msg_rc}")
            # --- FECHA REQUISIÇÃO SE OCORRER ERRO ---
            if msg_erro != "": 
                print(f"POSSUI ERRO | Mensagem ERRO: {msg_erro}")
        
                # --- PASTA ERROS ---            
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
                msg_rc = session.findById("wnd[0]/sbar").Text # Pega numero de requisição
                print(msg_rc)
                data = {'Chave': key, 'Mensagem Retorno': msg_rc}
                df_aux = pd.DataFrame(data, index=[index, index + 1])
                df_log = pd.concat([df_aux])
                print(df_log)

                # --- PASTA LOG ---
                user = dt.iloc[0]["Usuário"] # DEFININDO USUÁRIO  
                pasta = "LOG" # Nome da pasta LOG
                
                if not os.path.exists(pasta): # Cria pasta LOG se não existir
                    os.makedirs(pasta)
                    print(f"Pasta '{pasta}' criada com sucesso!")
                print("PASTA LOG EXISTE")
                log_user = f"LOG/LOG_{user}" # LOG POR USUÁRIO
                
                if not os.path.exists(log_user): # VERIFICAÇÃO E CRIAÇÃO DE PASTA LOG POR USUÁRIO
                    os.makedirs(log_user)
                    print(f"Pasta {log_user} criada com sucesso!")
                    print("LOG OK")

                nome_arquivo = file.split('/')[-1] # Limpando o nome do arquivo
                nome_arquivo = nome_arquivo.split('.')[0]

                df_log.to_excel(f"{log_user}/LOG_{hora_inicio}_{nome_arquivo}_{key}_.xlsx") # Definindo local e nome do arquivo
                print(file)

                # === REMOVE ARQUIVO ===
                #os.remove(file)
                #print(f"REMOVE ARQUIVO {file}")
                
                df_log = ""
                print(f"{log_user}/LOG_{hora_inicio}_{nome_arquivo}.xlsx")
                print(f"Mensagem retornada: {msg_rc} | CHAVE: {key} ")


        # Pegando informação de data e hora
        hora_inicio = datetime.datetime.now()
        hora_inicio = hora_inicio.strftime("%d-%m-%Y_%H-%M")

        #LEITURA DOS ARQUIVOS
        #lista_arquivos = get_files(dir=r"REQUISIÇÕES", return_dir= True) # Pasta Modelos de RC

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
            key = linha['Chave']
            tipo_rc = linha['Tipo RC'] 
            material = int(linha['Código'])
            texto = linha['Texto na RC']
            centro = linha['Centro']
            qtd = linha['Qtd']
            preco_av = linha['Preço Avaliação']
            deposito = int(linha['Depósito'])
            class_cont = linha['Class Cont']
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
            set_centro(centro, row = contador)
            set_qtd(qtd, row = contador)
            set_class(class_cont, row = contador)
            set_texto(texto)
            set_deposito(deposito, row = contador)
            
            # --- CENTRO DE CUSTO ---
            if tipo_rc == 'ZPA' or tipo_rc == 'ZRE' or tipo_rc == 'ZCM': # CENTRO DE CUSTO CIC
                print(f"\n RC {tipo_rc} CIC \n")
                set_cc_cic(cc, row = contador)
            else:
                print(f"\n RC {tipo_rc} \n") # CENTRO DE CUSTO NORMAL
                set_cc(cc, row = contador)
            
            set_preco_av(preco_av, row = contador)
            

            msg_erro = session.findById("wnd[0]/sbar").Text
            print(f"\nMSG_ERRO: {msg_erro}\n")

            msg_erro = session.findById("wnd[0]/sbar").Text
            print(f"\nMSG_ERRO: {msg_erro}\n")

            # --- CÓDIGO DE FORNECEDOR ---
            if pd.notnull(forn):
                set_fornecedor(forn, row = contador)
            else:
                pass
            
            msg_erro = session.findById("wnd[0]/sbar").Text
            print(f"\nMSG_ERRO: {msg_erro}\n")

            # --- NF ---
            if tipo_rc == "ZRE":
                set_nf(data_nf, data_ven, n_nf, row = contador)
            else:
                pass

            msg_erro = session.findById("wnd[0]/sbar").Text 
            print(f"\nMSG_ERRO: {msg_erro}\n")

            # --- CONTRATO | ITEM ---
            if pd.notnull(contrato and item):
                print("CONTRATO | ITEM")
                set_contrato(contrato, row = contador)
                set_item(item, row = contador)
            else:
                pass

            # MENSAGEM DE ERRO (barra inferior)
            msg_erro = session.findById("wnd[0]/sbar").Text
            print(f"\nMSG_ERRO: {msg_erro}\n")
            
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
                            #session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\gvm.indicadores\JSL SA\Grupo Vamos - Indicadores - Documentos\General\RC\anexos" # Local Anexos GVM.INDICADORES
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = anexos
                            #session.findById("wnd[1]/usr/ctxtDY_PATH").text = r"C:\Users\igor.gabriel\OneDrive - JSL SA\Documentos Compartilhados - Grupo Vamos - Indicadores\General\RC\anexos" # Local Anexos OneDrive
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
            
            time.sleep(1)
            
            contador += 1
            print(f"PROXIMA KEY: {proxima_key}\n")
            print(f"KEY: {key} {row}| TIPO RC: {tipo_rc} {row}| CODIGO: {material} {row}| CENTRO: {centro} {row}| QUANTIDADE: {qtd} {row}| PREÇO: {preco_av} {row}| DEPOSITO: {deposito} {row}| CLASS: {class_cont} {row}| CENTRO DE CUSTO: {cc} {row}")
            
            # --- VERIFICA SE É A ULTIMA LINHA ---
            if proxima_key != key and proxima_key is not None:
                print("---PROXIMA LINHA---")
                time.sleep(1)
                auto.hotkey('Ctrl','s')
                print("SALVO")
                time.sleep(2)
                auto.hotkey('F12')
                time.sleep(2)
                auto.hotkey('F6')
                time.sleep(3)
                auto.hotkey('tab')
                auto.hotkey('enter')
                print("ABRIU NOVA REQUISIÇÃO")
                contador = 0
                retorno()

            elif proxima_key is None:
                print("--- FECHANDO RC ---")
                time.sleep(3)
                auto.hotkey('Ctrl','s')
                time.sleep(3)
                contador = 0
                print("SALVOU")
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
                print("CTRL + S")
                time.sleep(3)
                contador = 0
                auto.hotkey('F12')
                print("F12")
                time.sleep(1)
                auto.hotkey('F6')
                print("F6")
                time.sleep(1)
                auto.hotkey('tab')
                print("TAB")
                auto.hotkey('enter')
                print("ENTER")
                retorno()

                # --- VERIFICAÇÃO SALVAMENTO COM ANEXO ---
                if anexo != pd.notnull(anexo) and msg_erro == "": 
                    time.sleep(1)
                    print(f"CHAVE: {key} (SALVO COM ANEXO)\n")  
                    #auto.press('tab', presses=4)
                    #auto.hotkey('enter')
                    #auto.press('tab', presses=4)
                    #auto.hotkey('enter')
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
print("... FIM ...")
