# -*- encoding: utf-8 -*-
# pip install --upgrade virtualenv google-api-python-client google-auth-httplib2 google-auth-oauthlib gspread oauth2client
# https://github.com/faustostangler/b3-bovespa/edit/master/allinone.py

# selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

# datetime
from datetime import datetime
from datetime import timedelta
import time

# google
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import re


# back-to-basics
def user_defined_variables():
    try:
        # browser path
        global base_azevedo
        global base_note
        global base_hbo
        global chrome
        global firefox

        base_azevedo = 'C:/Users/faust/PycharmProjects/'
        base_note = 'C:/Users/Fausto Stangler/PycharmProjects/b3A-companies/'
        base_hbo = 'C:/Users/Fausto/PycharmProjects/b3A-companies/'
        chrome = 'chromedriver.exe'
        firefox = 'firefox.exe'

        # main page and sheet and cvm base URLs
        global main_url
        global sheet_url
        global cvm_url
        main_url = 'http://bvmf.bmfbovespa.com.br/cias-listadas/empresas-listadas/BuscaEmpresaListada.aspx?idioma=pt-br'
        sheet_url = 'https://docs.google.com/spreadsheets/d/'
        cvm_url = 'http://bvmf.bmfbovespa.com.br/pt-br/mercados/acoes/empresas/ExecutaAcaoConsultaInfoEmp.asp?CodCVM='

        # ID of Google Sheet to store list of companies
        global index_sheet
        index_sheet = '1AZzuxXbhmDbp5hFOcvYPZK0zYh91uofX7Ry0DcIvIdk'

        # ID of Google Sheet Template for new company sheets
        global report_sheet
        global fundament_sheet
        global default_sheets
        report_sheet = '1Bn9t20r3czSiqK4S76bd7oL1UCthmDnPw32GMhU1k1s' # MODELO REPORT
        fundament_sheet = '1R5YOiJOWGhjgvEQdtst_PqrPMeGLV1qe9imjA4lk6mg' # MODELO FUNDAMENTOS

        default_sheets = ['INDEX']  # , 'F-WEB', 'F', 'Glossário', 'blasterlista', 'quotes']

        # google API factor of true
        global google
        google = False

        # google_sheet API authorization - see https://console.developers.google.com/
        global CLIENT_SECRET_FILE
        global ACCOUNT_SECRET_FILE
        global SHEET_SCOPE
        CLIENT_SECRET_FILE = 'client_credentials.json'
        ACCOUNT_SECRET_FILE = 'account_credentials.json'
        SHEET_SCOPE = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

        # google drive API authorization - see https://console.developers.google.com/
        global CLIENT_EMAIL
        global DRIVE_SCOPE
        CLIENT_EMAIL = 'dre-empresas-listadas-bot@dre-empresas-listadas-b3.iam.gserviceaccount.com'
        DRIVE_SCOPE = ['https://www.googleapis.com/auth/drive']

        # pre-defined quantities
        global batch_companies
        batch_companies = 500

        global batch_reports
        batch_reports = 120

        # reports from b3
        global parts
        global url
        parts = [['Dados', 'Dados da Empresa', 'Composição do Capital'],
                 ['DRE', 'DFs Individuais', 'Balanço Patrimonial Ativo'],
                 ['DRE', 'DFs Individuais', 'Balanço Patrimonial Passivo'],
                 ['DRE', 'DFs Individuais', 'Demonstração do Resultado'],
                 ['DRE', 'DFs Individuais', 'Demonstração do Resultado Abrangente'],
                 ['DRE', 'DFs Individuais', 'Demonstração do Fluxo de Caixa'],
                 ['DRE', 'DFs Individuais', 'Demonstração de Valor Adicionado'],
                 ['DRE', 'DFs Consolidadas', 'Balanço Patrimonial Ativo'],
                 ['DRE', 'DFs Consolidadas', 'Balanço Patrimonial Passivo'],
                 ['DRE', 'DFs Consolidadas', 'Demonstração do Resultado'],
                 ['DRE', 'DFs Consolidadas', 'Demonstração do Resultado Abrangente'],
                 ['DRE', 'DFs Consolidadas', 'Demonstração do Fluxo de Caixa'],
                 ['DRE', 'DFs Consolidadas', 'Demonstração de Valor Adicionado']]
        url = ''

        global DRE
        DRE = ''

    except Exception as e:
        restart(e, __name__)
def startEngine():
    try:
        global browser
        global wait

        print('BROWSER start')
        # load browser and general parameters
        try:
            browser = webdriver.Chrome(executable_path=base_azevedo + chrome)
        except:
            pass
        try:
            browser = webdriver.Chrome(executable_path=base_note + chrome)
        except:
            pass
        try:
            browser = webdriver.Chrome(executable_path=base_hbo + chrome)
        except:
            pass
        wait = WebDriverWait(browser, 60)
        browser.minimize_window()

        print('...done')
    except Exception as e:
        restart(e, __name__)
def googleAPI():
    try:
        global gdrive

        print('GOOGLE API authorization')

        # google drive authorization
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, DRIVE_SCOPE)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)
        gdrive = build('drive', 'v3', credentials=creds)

        # google sheet authorization
        global gsheet
        credentials = ServiceAccountCredentials.from_json_keyfile_name(ACCOUNT_SECRET_FILE, SHEET_SCOPE)
        gsheet = gspread.authorize(credentials)

        loadSheets()

        print('...done')
        return True
    except Exception as e:
        print('google API error... restarting')
        google = googleAPI()
def loadSheets():
    try:
        # sheet worksheets
        global bovespa
        global bovespa_listagem
        global bovespa_log

        bovespa = gsheet.open_by_key(index_sheet)
        bovespa_listagem = bovespa.worksheet('listagem')
        bovespa_log = bovespa.worksheet('log')
    except Exception as e:
        restart(e, __name__)
def loadCompanyReportsSheets(company):
    try:
        global report_sheet_index
        global report_sheet_reports

        company_sheet_report = gsheet.open_by_key(company [col ['REPORTS']].replace(sheet_url, ''))
        report_sheet_index = company_sheet_report.worksheet('index')
        report_sheet_reports = company_sheet_report.worksheet('reports')

    except Exception as e:
        restart(e, __name__)
def loadCompanyFundamentosSheets(company):
    try:
        global fundamentos_sheet_index
        global fundamentos_sheet_reports

        company_sheet_fundamentos = gsheet.open_by_key(company [col ['FUNDAMENTOS']].replace(sheet_url, ''))
        fundamentos_sheet_index = company_sheet_fundamentos.worksheet('index')
        fundamentos_sheet_reports = company_sheet_fundamentos.worksheet('reports')

    except Exception as e:
        restart(e, __name__)
def list_unique(li1, li2):
    try:
        li3 = []
        [li3.append(i) for i in li1 + li2 if i not in li3]
        return li3
    except Exception as e:
        restart(e, __name__)
def list_remove_extra(li1, li2):
    try:
        li3 = [i for i in li1 if i not in li2]
        return li3
    except Exception as e:
        restart(e, __name__)
def list_difference(li1, li2):
    try:
        li_dif = [i for i in li1 + li2 if i not in li1 or i not in li2]
        return li_dif
    except Exception as e:
        restart(e, __name__)
def list_intersection(li1, li2):
    try:
        li3 = [value for value in li1 if value in li2]
        return li3
    except Exception as e:
        restart(e, __name__)
def sheetCol(n):
    try:
        string = ''
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string
    except Exception as e:
        restart(e, __name__)
def sheetRange(start_row, start_col, data):
    try:
        range1 = str(sheetCol(start_col)) + str(start_row)
        range2 = str(sheetCol(start_col + len(data [0]) - 1)) + str(len(data) + start_row - 1)
        sheet_range = range1 + ':' + range2
        return sheet_range
    except Exception as e:
        restart(e, __name__)
def sheetColumns(sheet_columns):
    try:
        global col
        col = {k: v for v, k in enumerate(sheet_columns)}
        # print(col ['EMPRESA'])
        return sheet_columns
    except Exception as e:
        restart(e, __name__)
def sheetCompany():
    try:
        global google
        if google != True:
            google = googleAPI()
        loadSheets()

        print('LOAD sheet of companies')

        # get planilha all values
        try:
            sheet_companies = []
            sheet_companies = bovespa_listagem.get_all_values()
            sheet_columns = sheetColumns(sheet_companies.pop(0))
        except:
            print('sheet_of_companies =', sheet_companies)
        print('...done')
        return sheet_companies
    except Exception as e:
        restart(e, __name__)
def b3Company():
    try:
        b3_companies = []
        # b3_companies = [['16284', '524 PARTICIP', '524 PARTICIPACOES S.A.', 'MB'],
        #                 ['21725', 'ADVANCED-DH', 'ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.', ' '],
        #                 ['18970', 'AES TIETE E', 'AES TIETE ENERGIA SA', 'N2'],
        #                 ['22179', 'AFLUENTE T', 'AFLUENTE TRANSMISSÃO DE ENERGIA ELÉTRICA S/A', ' ']]
        # b3_companies = [['16284', '524 PARTICIP', '524 PARTICIPACOES S.A.', 'MB'],
        #                 ['21725', 'ADVANCED-DH', 'ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.', ' '],
        #                 ['18970', 'AES TIETE E', 'AES TIETE ENERGIA SA', 'N2'],
        #                 ['22179', 'AFLUENTE T', 'AFLUENTE TRANSMISSÃO DE ENERGIA ELÉTRICA S/A', ' '],
        #                 ['16705', 'ALEF S/A', 'ALEF S.A.', 'MB'], ['9954', 'ALFA HOLDING', 'ALFA HOLDINGS S.A.', ' '],
        #                 ['21032', 'ALGAR TELEC', 'ALGAR TELECOM S/A', ' '],
        #                 ['22357', 'ALIANSCSONAE', 'ALIANSCE SONAE SHOPPING CENTERS S.A.', 'NM'],
        #                 ['24953', 'ESTAPAR', 'ALLPARK EMPREENDIMENTOS PARTICIPACOES SERVICOS S.A', 'NM'],
        #                 ['10456', 'ALPARGATAS', 'ALPARGATAS S.A.', 'N1'],
        #                 ['22217', 'ALPER S.A.', 'ALPER CONSULTORIA E CORRETORA DE SEGUROS S.A.', 'NM'],
        #                 ['18066', 'ALTERE SEC', 'ALTERE SECURITIZADORA S.A.', ' '],
        #                 ['21490', 'ALUPAR', 'ALUPAR INVESTIMENTO S/A', 'N2'], ['23264', 'AMBEV S/A', 'AMBEV S.A.', ' '],
        #                 ['3050', 'AMPLA ENERG', 'AMPLA ENERGIA E SERVICOS S.A.', ' '],
        #                 ['23248', 'ANIMA', 'ANIMA HOLDING S.A.', 'NM'],
        #                 ['22349', 'AREZZO CO', 'AREZZO INDÚSTRIA E COMÉRCIO S.A.', 'NM'],
        #                 ['24171', 'CARREFOUR BR', 'ATACADÃO S.A.', 'NM'],
        #                 ['19100', 'ATMASA', 'ATMA PARTICIPAÇÕES S.A.', 'NM'],
        #                 ['15423', 'ATOMPAR', 'ATOM EMPREENDIMENTOS E PARTICIPAÇÕES S.A.', ' '],
        #                 ['11975', 'AZEVEDO', 'AZEVEDO E TRAVASSOS S.A.', ' '], ['24112', 'AZUL', 'AZUL S.A.', 'N2'],
        #                 ['20990', 'B2W DIGITAL', 'B2W - COMPANHIA DIGITAL', 'NM'],
        #                 ['21610', 'B3', 'B3 S.A. - BRASIL. BOLSA. BALCÃO', 'NM'],
        #                 ['701', 'BAHEMA', 'BAHEMA EDUCAÇÃO S.A.', 'MA'], ['24600', 'BANCO BMG', 'BANCO BMG S.A.', 'N1'],
        #                 ['24406', 'BANCO INTER', 'BANCO INTER S.A.', 'N2'],
        #                 ['1155', 'BANESTES', 'BANESTES S.A. - BCO EST ESPIRITO SANTO', ' '],
        #                 ['1520', 'BARDELLA', 'BARDELLA S.A. INDUSTRIAS MECANICAS', ' '],
        #                 ['15458', 'BATTISTELLA', 'BATTISTELLA ADM PARTICIPACOES S.A.', ' '],
        #                 ['1562', 'BAUMER', 'BAUMER S.A.', ' '],
        #                 ['23159', 'BBSEGURIDADE', 'BB SEGURIDADE PARTICIPAÇÕES S.A.', 'NM'],
        #                 ['24660', 'BBMLOGISTICA', 'BBM LOGISTICA S.A.', 'MA'],
        #                 ['20958', 'ABC BRASIL', 'BCO ABC BRASIL S.A.', 'N2'],
        #                 ['1384', 'ALFA INVEST', 'BCO ALFA DE INVESTIMENTO S.A.', ' '],
        #                 ['922', 'AMAZONIA', 'BCO AMAZONIA S.A.', ' '], ['906', 'BRADESCO', 'BCO BRADESCO S.A.', 'N1'],
        #                 ['1023', 'BRASIL', 'BCO BRASIL S.A.', 'NM'],
        #                 ['22616', 'BTGP BANCO', 'BCO BTG PACTUAL S.A.', 'N2'],
        #                 ['1120', 'BANESE', 'BCO ESTADO DE SERGIPE S.A. - BANESE', ' '],
        #                 ['1171', 'BANPARA', 'BCO ESTADO DO PARA S.A.', ' '],
        #                 ['1210', 'BANRISUL', 'BCO ESTADO DO RIO GRANDE DO SUL S.A.', 'N1'],
        #                 ['20885', 'INDUSVAL', 'BCO INDUSVAL S.A.', 'N2'],
        #                 ['1309', 'MERC INVEST', 'BCO MERCANTIL DE INVESTIMENTOS S.A.', ' '],
        #                 ['1325', 'MERC BRASIL', 'BCO MERCANTIL DO BRASIL S.A.', ' '],
        #                 ['1228', 'NORD BRASIL', 'BCO NORDESTE DO BRASIL S.A.', ' '],
        #                 ['21199', 'BANCO PAN', 'BCO PAN S.A.', 'N1'], ['20567', 'PINE', 'BCO PINE S.A.', 'N2'],
        #                 ['20532', 'SANTANDER BR', 'BCO SANTANDER (BRASIL) S.A.', ' '],
        #                 ['19747', 'BETA SECURIT', 'BETA SECURITIZADORA S.A.', ' '],
        #                 ['17884', 'BETAPART', 'BETAPART PARTICIPACOES S.A.', 'MB'],
        #                 ['1694', 'BIC MONARK', 'BICICLETAS MONARK S.A.', ' '], ['19305', 'BIOMM', 'BIOMM S.A.', 'MA'],
        #                 ['22845', 'BIOSEV', 'BIOSEV S.A.', 'NM'],
        #                 ['80179', 'BIOTOSCANA', 'BIOTOSCANA INVESTMENTS S.A.', 'DR3'],
        #                 ['24317', 'BK BRASIL', 'BK BRASIL OPERAÇÃO E ASSESSORIA A RESTAURANTES SA', 'NM'],
        #                 ['16772', 'BNDESPAR', 'BNDES PARTICIPACOES S.A. - BNDESPAR', 'MB'],
        #                 ['12190', 'BOMBRIL', 'BOMBRIL S.A.', ' '],
        #                 ['19909', 'BR MALLS PAR', 'BR MALLS PARTICIPACOES S.A.', 'NM'],
        #                 ['19925', 'BR PROPERT', 'BR PROPERTIES S.A.', 'NM'],
        #                 ['19640', 'BRADESCO LSG', 'BRADESCO LEASING S.A. ARREND MERCANTIL', ' '],
        #                 ['18724', 'BRADESPAR', 'BRADESPAR S.A.', 'N1'],
        #                 ['21180', 'BR BROKERS', 'BRASIL BROKERS PARTICIPACOES S.A.', 'NM'],
        #                 ['20036', 'BRASILAGRO', 'BRASILAGRO - CIA BRAS DE PROP AGRICOLAS', 'NM'],
        #                 ['4820', 'BRASKEM', 'BRASKEM S.A.', 'N1'],
        #                 ['19720', 'BRAZIL REALT', 'BRAZIL REALTY CIA SECURIT. CRÉD. IMOBILIÁRIOS', ' '],
        #                 ['17922', 'BRAZILIAN FR', 'BRAZILIAN FINANCE E REAL ESTATE S.A.', ' '],
        #                 ['18759', 'BRAZILIAN SC', 'BRAZILIAN SECURITIES CIA SECURITIZACAO', ' '],
        #                 ['14206', 'BRB BANCO', 'BRB BCO DE BRASILIA S.A.', ' '],
        #                 ['20672', 'BRC SECURIT', 'BRC SECURITIZADORA S.A.', ' '], ['16292', 'BRF SA', 'BRF S.A.', 'NM'],
        #                 ['19984', 'BRPR 56 SEC', 'BRPR 56 SECURITIZADORA CRED IMOB S.A.', ' '],
        #                 ['23817', 'BRQ', 'BRQ SOLUCOES EM INFORMATICA S.A.', 'MA'],
        #                 ['20133', 'BV LEASING', 'BV LEASING - ARRENDAMENTO MERCANTIL S.A.', ' '],
        #                 ['19119', 'CABINDA PART', 'CABINDA PARTICIPACOES S.A.', 'MB'],
        #                 ['22683', 'CACHOEIRA', 'CACHOEIRA PAULISTA TRANSMISSORA ENERGIA S.A.', 'MB'],
        #                 ['19135', 'CACONDE PART', 'CACONDE PARTICIPACOES S.A.', 'MB'],
        #                 ['2100', 'CAMBUCI', 'CAMBUCI S.A.', ' '], ['24228', 'CAMIL', 'CAMIL ALIMENTOS S.A.', 'NM'],
        #                 ['17493', 'CAPITALPART', 'CAPITALPART PARTICIPACOES S.A.', 'MB'],
        #                 ['18821', 'CCR SA', 'CCR S.A.', 'NM'], ['24848', 'CEA MODAS', 'CEA MODAS S.A.', 'NM'],
        #                 ['13854', 'CEMEPE', 'CEMEPE INVESTIMENTOS S.A.', ' '],
        #                 ['20303', 'CEMIG DIST', 'CEMIG DISTRIBUICAO S.A.', ' '],
        #                 ['20320', 'CEMIG GT', 'CEMIG GERACAO E TRANSMISSAO S.A.', ' '],
        #                 ['2437', 'ELETROBRAS', 'CENTRAIS ELET BRAS S.A. - ELETROBRAS', 'N1'],
        #                 ['2461', 'CELESC', 'CENTRAIS ELET DE SANTA CATARINA S.A.', 'N2'],
        #                 ['24058', 'ALLIAR', 'CENTRO DE IMAGEM DIAGNOSTICOS S.A.', 'NM'],
        #                 ['2577', 'CESP', 'CESP - CIA ENERGETICA DE SAO PAULO', 'N1'],
        #                 ['14826', 'P.ACUCAR-CBD', 'CIA BRASILEIRA DE DISTRIBUICAO', 'NM'],
        #                 ['16861', 'CASAN', 'CIA CATARINENSE DE AGUAS E SANEAM.-CASAN', ' '],
        #                 ['21393', 'CELGPAR', 'CIA CELG DE PARTICIPACOES - CELGPAR', ' '],
        #                 ['16616', 'CEG', 'CIA DISTRIB DE GAS DO RIO DE JANEIRO-CEG', ' '],
        #                 ['14524', 'COELBA', 'CIA ELETRICIDADE EST. DA BAHIA - COELBA', ' '],
        #                 ['14451', 'CEB', 'CIA ENERGETICA DE BRASILIA', ' '],
        #                 ['2453', 'CEMIG', 'CIA ENERGETICA DE MINAS GERAIS - CEMIG', 'N1'],
        #                 ['14362', 'CELPE', 'CIA ENERGETICA DE PERNAMBUCO - CELPE', ' '],
        #                 ['14869', 'COELCE', 'CIA ENERGETICA DO CEARA - COELCE', ' '],
        #                 ['18139', 'COSERN', 'CIA ENERGETICA DO RIO GDE NORTE - COSERN', ' '],
        #                 ['20648', 'CEEE-D', 'CIA ESTADUAL DE DISTRIB ENER ELET-CEEE-D', 'N1'],
        #                 ['3204', 'CEEE-GT', 'CIA ESTADUAL GER.TRANS.ENER.ELET-CEEE-GT', 'N1'],
        #                 ['3069', 'FERBASA', 'CIA FERRO LIGAS DA BAHIA - FERBASA', 'N1'],
        #                 ['3077', 'CEDRO', 'CIA FIACAO TECIDOS CEDRO CACHOEIRA', 'N1'],
        #                 ['15636', 'COMGAS', 'CIA GAS DE SAO PAULO - COMGAS', ' '],
        #                 ['3298', 'HABITASUL', 'CIA HABITASUL DE PARTICIPACOES', ' '],
        #                 ['14761', 'CIA HERING', 'CIA HERING', 'NM'],
        #                 ['3395', 'IND CATAGUAS', 'CIA INDUSTRIAL CATAGUASES', ' '],
        #                 ['22691', 'LOCAMERICA', 'CIA LOCAÇÃO DAS AMÉRICAS', 'NM'],
        #                 ['3654', 'MELHOR SP', 'CIA MELHORAMENTOS DE SAO PAULO', ' '],
        #                 ['14311', 'COPEL', 'CIA PARANAENSE DE ENERGIA - COPEL', 'N1'],
        #                 ['18708', 'PAR AL BAHIA', 'CIA PARTICIPACOES ALIANCA DA BAHIA', ' '],
        #                 ['3824', 'PAUL F LUZ', 'CIA PAULISTA DE FORCA E LUZ', ' '],
        #                 ['19275', 'CPFL PIRATIN', 'CIA PIRATININGA DE FORCA E LUZ', ' '],
        #                 ['14443', 'SABESP', 'CIA SANEAMENTO BASICO EST SAO PAULO', 'NM'],
        #                 ['19445', 'COPASA', 'CIA SANEAMENTO DE MINAS GERAIS-COPASA MG', 'NM'],
        #                 ['18627', 'SANEPAR', 'CIA SANEAMENTO DO PARANA - SANEPAR', 'N2'],
        #                 ['3115', 'SEG AL BAHIA', 'CIA SEGUROS ALIANCA DA BAHIA', ' '],
        #                 ['4030', 'SID NACIONAL', 'CIA SIDERURGICA NACIONAL', ' '],
        #                 ['3158', 'COTEMINAS', 'CIA TECIDOS NORTE DE MINAS COTEMINAS', ' '],
        #                 ['4081', 'SANTANENSE', 'CIA TECIDOS SANTANENSE', ' '],
        #                 ['18287', 'CIBRASEC', 'CIBRASEC - COMPANHIA BRASILEIRA DE SECURITIZACAO', ' '],
        #                 ['21733', 'CIELO', 'CIELO S.A.', 'NM'], ['14818', 'CIMS', 'CIMS S.A.', ' '],
        #                 ['23965', 'CINESYSTEM', 'CINESYSTEM S.A.', 'MA'],
        #                 ['17973', 'COGNA ON', 'COGNA EDUCAÇÃO S.A.', 'NM'],
        #                 ['22268', 'CONC RAPOSO', 'CONC AUTO RAPOSO TAVARES S.A.', ' '],
        #                 ['23515', 'GRUAIRPORT', 'CONC DO AEROPORTO INTERNACIONAL DE GUARULHOS S.A.', 'MB'],
        #                 ['20397', 'ECOVIAS', 'CONC ECOVIAS IMIGRANTES S.A.', ' '],
        #                 ['19208', 'CONC RIO TER', 'CONC RIO-TERESOPOLIS S.A.', 'MB'],
        #                 ['22411', 'ECOPISTAS', 'CONC ROD AYRTON SENNA E CARV PINTO S.A.-ECOPISTAS', ' '],
        #                 ['21024', 'VIAOESTE', 'CONC ROD.OESTE SP VIAOESTE S.A', ' '],
        #                 ['22721', 'ROD TIETE', 'CONC RODOVIAS DO TIETÊ S.A.', ' '],
        #                 ['22071', 'RT BANDEIRAS', 'CONC ROTA DAS BANDEIRAS S.A.', ' '],
        #                 ['20192', 'AUTOBAN', 'CONC SIST ANHANG-BANDEIRANT S.A. AUTOBAN', ' '],
        #                 ['4693', 'ODERICH', 'CONSERVAS ODERICH S.A.', ' '],
        #                 ['4707', 'ALFA CONSORC', 'CONSORCIO ALFA DE ADMINISTRACAO S.A.', ' '],
        #                 ['4723', 'CONST A LIND', 'CONSTRUTORA ADOLPHO LINDENBERG S.A.', ' '],
        #                 ['21148', 'TENDA', 'CONSTRUTORA TENDA S.A.', 'NM'],
        #                 ['4863', 'COR RIBEIRO', 'CORREA RIBEIRO S.A. COMERCIO E INDUSTRIA', ' '],
        #                 ['23485', 'COSAN LOG', 'COSAN LOGISTICA S.A.', 'NM'], ['19836', 'COSAN', 'COSAN S.A.', 'NM'],
        #                 ['18660', 'CPFL ENERGIA', 'CPFL ENERGIA S.A.', 'NM'],
        #                 ['20540', 'CPFL RENOVAV', 'CPFL ENERGIAS RENOVÁVEIS S.A.', 'NM'],
        #                 ['18953', 'CPFL GERACAO', 'CPFL GERACAO DE ENERGIA S.A.', ' '],
        #                 ['20630', 'CR2', 'CR2 EMPREENDIMENTOS IMOBILIARIOS S.A.', ' '],
        #                 ['20044', 'CSU CARDSYST', 'CSU CARDSYSTEM S.A.', 'NM'],
        #                 ['23981', 'CTC S.A.', 'CTC - CENTRO DE TECNOLOGIA CANAVIEIRA S.A.', 'MA'],
        #                 ['18376', 'TRAN PAULIST', 'CTEEP - CIA TRANSMISSÃO ENERGIA ELÉTRICA PAULISTA', 'N1'],
        #                 ['23310', 'CVC BRASIL', 'CVC BRASIL OPERADORA E AGÊNCIA DE VIAGENS S.A.', 'NM'],
        #                 ['14460', 'CYRELA REALT', 'CYRELA BRAZIL REALTY S.A.EMPREEND E PART', 'NM'],
        #                 ['21040', 'CYRE COM-CCP', 'CYRELA COMMERCIAL PROPERT S.A. EMPR PART', 'NM'],
        #                 ['19623', 'DASA', 'DIAGNOSTICOS DA AMERICA S.A.', ' '],
        #                 ['14214', 'DIBENS LSG', 'DIBENS LEASING S.A. - ARREND.MERCANTIL', ' '],
        #                 ['9342', 'DIMED', 'DIMED S.A. DISTRIBUIDORA DE MEDICAMENTOS', ' '],
        #                 ['21350', 'DIRECIONAL', 'DIRECIONAL ENGENHARIA S.A.', 'NM'],
        #                 ['5207', 'DOHLER', 'DOHLER S.A.', ' '], ['23493', 'DOMMO', 'DOMMO ENERGIA S.A.', ' '],
        #                 ['18597', 'DTCOM-DIRECT', 'DTCOM - DIRECT TO COMPANY S.A.', ' '],
        #                 ['21091', 'DURATEX', 'DURATEX S.A.', 'NM'],
        #                 ['21741', 'ECO SEC AGRO', 'ECO SECURITIZADORA DIREITOS CRED AGRONEGÓCIO S.A.', 'MB'],
        #                 ['21903', 'ECON', 'ECORODOVIAS CONCESSÕES E SERVIÇOS S.A.', ' '],
        #                 ['19453', 'ECORODOVIAS', 'ECORODOVIAS INFRAESTRUTURA E LOGÍSTICA S.A.', 'NM'],
        #                 ['19763', 'ENERGIAS BR', 'EDP - ENERGIAS DO BRASIL S.A.', 'NM'],
        #                 ['15342', 'ESCELSA', 'EDP ESPIRITO SANTO DISTRIBUIÇÃO DE ENERGIA S.A.', ' '],
        #                 ['16985', 'EBE', 'EDP SÃO PAULO DISTRIBUIÇÃO DE ENERGIA S.A.', ' '],
        #                 ['5380', 'ACO ALTONA', 'ELECTRO ACO ALTONA S.A.', ' '],
        #                 ['4359', 'ELEKEIROZ', 'ELEKEIROZ S.A.', ' '], ['17485', 'ELEKTRO', 'ELEKTRO REDES S.A.', ' '],
        #                 ['15784', 'ELETROPAR', 'ELETROBRÁS PARTICIPAÇÕES S.A. - ELETROPAR', ' '],
        #                 ['14176', 'ELETROPAULO', 'ELETROPAULO METROP. ELET. SAO PAULO S.A.', ' '],
        #                 ['16993', 'EMAE', 'EMAE - EMPRESA METROP.AGUAS ENERGIA S.A.', ' '],
        #                 ['20087', 'EMBRAER', 'EMBRAER S.A.', 'NM'],
        #                 ['19011', 'ECONORTE', 'EMPRESA CONC RODOV DO NORTE S.A.ECONORTE', ' '],
        #                 ['16497', 'ENCORPAR', 'EMPRESA NAC COM REDITO PART S.A.ENCORPAR', ' '],
        #                 ['22365', 'ENAUTA PART', 'ENAUTA PARTICIPAÇÕES S.A.', 'NM'],
        #                 ['5576', 'ENERSUL', 'ENERGISA MATO GROSSO DO SUL - DIST DE ENERGIA S.A.', ' '],
        #                 ['14605', 'ENERGISA MT', 'ENERGISA MATO GROSSO-DISTRIBUIDORA DE ENERGIA S/A', ' '],
        #                 ['15253', 'ENERGISA', 'ENERGISA S.A.', 'N2'], ['21237', 'ENEVA', 'ENEVA S.A', 'NM'],
        #                 ['17329', 'ENGIE BRASIL', 'ENGIE BRASIL ENERGIA S.A.', 'NM'],
        #                 ['20010', 'EQUATORIAL', 'EQUATORIAL ENERGIA S.A.', 'NM'],
        #                 ['16608', 'EQTLMARANHAO', 'EQUATORIAL MARANHÃO DISTRIBUIDORA DE ENERGIA S.A.', 'MB'],
        #                 ['18309', 'EQTL PARA', 'EQUATORIAL PARA DISTRIBUIDORA DE ENERGIA S.A.', ' '],
        #                 ['5762', 'ETERNIT', 'ETERNIT S.A.', 'NM'],
        #                 ['5770', 'EUCATEX', 'EUCATEX S.A. INDUSTRIA E COMERCIO', 'N1'],
        #                 ['20524', 'EVEN', 'EVEN CONSTRUTORA E INCORPORADORA S.A.', 'NM'],
        #                 ['1570', 'EXCELSIOR', 'EXCELSIOR ALIMENTOS S.A.', ' '],
        #                 ['20770', 'EZTEC', 'EZ TEC EMPREEND. E PARTICIPACOES S.A.', 'NM'],
        #                 ['22977', 'FGENERGIA', 'FERREIRA GOMES ENERGIA S.A.', ' '],
        #                 ['15369', 'FER C ATLANT', 'FERROVIA CENTRO-ATLANTICA S.A.', ' '],
        #                 ['20621', 'FER HERINGER', 'FERTILIZANTES HERINGER S.A.', 'NM'],
        #                 ['3891', 'ALFA FINANC', 'FINANCEIRA ALFA S.A.- CRED FINANC E INVS', ' '],
        #                 ['6076', 'FINANSINOS', 'FINANSINOS S.A.- CREDITO FINANC E INVEST', ' '],
        #                 ['21881', 'FLEURY', 'FLEURY S.A.', 'NM'],
        #                 ['24350', 'FLEX S/A', 'FLEX GESTÃO DE RELACIONAMENTOS S.A.', 'MA'],
        #                 ['6211', 'FRAS-LE', 'FRAS-LE S.A.', 'N1'], ['16101', 'GAFISA', 'GAFISA S.A.', 'NM'],
        #                 ['22764', 'GAIA AGRO', 'GAIA AGRO SECURITIZADORA S.A.', ' '],
        #                 ['20222', 'GAIA SECURIT', 'GAIA SECURITIZADORA S.A.', 'MB'],
        #                 ['17965', 'GAMA PART', 'GAMA PARTICIPACOES S.A.', 'MB'],
        #                 ['21008', 'GENERALSHOPP', 'GENERAL SHOPPING E OUTLETS DO BRASIL S.A.', 'NM'],
        #                 ['3980', 'GERDAU', 'GERDAU S.A.', 'N1'],
        #                 ['19569', 'GOL', 'GOL LINHAS AEREAS INTELIGENTES S.A.', 'N2'],
        #                 ['80020', 'GP INVEST', 'GP INVESTMENTS. LTD.', 'DR3'],
        #                 ['16632', 'GPC PART', 'GPC PARTICIPACOES S.A.', ' '],
        #                 ['4537', 'GRAZZIOTIN', 'GRAZZIOTIN S.A.', ' '], ['19615', 'GRENDENE', 'GRENDENE S.A.', 'NM'],
        #                 ['24694', 'CENTAURO', 'GRUPO SBF SA', 'NM'],
        #                 ['4669', 'GUARARAPES', 'GUARARAPES CONFECCOES S.A.', ' '],
        #                 ['13366', 'HAGA S/A', 'HAGA S.A. INDUSTRIA E COMERCIO', ' '],
        #                 ['24392', 'HAPVIDA', 'HAPVIDA PARTICIPACOES E INVESTIMENTOS SA', 'NM'],
        #                 ['20877', 'HELBOR', 'HELBOR EMPREENDIMENTOS S.A.', 'NM'],
        #                 ['6629', 'HERCULES', 'HERCULES S.A. FABRICA DE TALHERES', ' '],
        #                 ['6700', 'HOTEIS OTHON', 'HOTEIS OTHON S.A.', ' '], ['21431', 'HYPERA', 'HYPERA S.A.', 'NM'],
        #                 ['18414', 'IDEIASNET', 'IDEIASNET S.A.', ' '], ['6815', 'IGB S/A', 'IGB ELETRÔNICA S/A', ' '],
        #                 ['23175', 'IGUA SA', 'IGUA SANEAMENTO S.A.', 'MA'],
        #                 ['20494', 'IGUATEMI', 'IGUATEMI EMPRESA DE SHOPPING CENTERS S.A', 'NM'],
        #                 ['12319', 'J B DUARTE', 'INDUSTRIAS J B DUARTE S.A.', ' '],
        #                 ['7510', 'INDS ROMI', 'INDUSTRIAS ROMI S.A.', 'NM'],
        #                 ['7595', 'INEPAR', 'INEPAR S.A. INDUSTRIA E CONSTRUCOES', ' '],
        #                 ['17558', 'SELECTPART', 'INNCORP S.A.', 'MB'],
        #                 ['24090', 'IHPARDINI', 'INSTITUTO HERMES PARDINI S.A.', 'NM'],
        #                 ['24279', 'INTER SA', 'INTER CONSTRUTORA E INCORPORADORA S.A.', 'MA'],
        #                 ['23574', 'IMC S/A', 'INTERNATIONAL MEAL COMPANY ALIMENTACAO S.A.', 'NM'],
        #                 ['6041', 'INVEST BEMGE', 'INVESTIMENTOS BEMGE S.A.', ' '],
        #                 ['18775', 'INVEPAR', 'INVESTIMENTOS E PARTICIP. EM INFRA S.A. - INVEPAR', 'MB'],
        #                 ['11932', 'IOCHP-MAXION', 'IOCHPE MAXION S.A.', 'NM'],
        #                 ['2429', 'IRANI', 'IRANI PAPEL E EMBALAGEM S.A.', ' '],
        #                 ['24180', 'IRBBRASIL RE', 'IRB - BRASIL RESSEGUROS S.A.', 'NM'],
        #                 ['19364', 'ITAPEBI', 'ITAPEBI GERACAO DE ENERGIA S.A.', ' '],
        #                 ['19348', 'ITAUUNIBANCO', 'ITAU UNIBANCO HOLDING S.A.', 'N1'],
        #                 ['7617', 'ITAUSA', 'ITAUSA INVESTIMENTOS ITAU S.A.', 'N1'],
        #                 ['21156', 'J.MACEDO', 'J. MACEDO S.A.', ' '], ['20575', 'JBS', 'JBS S.A.', 'NM'],
        #                 ['8672', 'JEREISSATI', 'JEREISSATI PARTICIPACOES S.A.', ' '],
        #                 ['20605', 'JHSF PART', 'JHSF PARTICIPACOES S.A.', 'NM'],
        #                 ['7811', 'JOAO FORTES', 'JOAO FORTES ENGENHARIA S.A.', ' '],
        #                 ['13285', 'JOSAPAR', 'JOSAPAR-JOAQUIM OLIVEIRA S.A. - PARTICIP', ' '],
        #                 ['22020', 'JSL', 'JSL S.A.', 'NM'], ['4146', 'KARSTEN', 'KARSTEN S.A.', ' '],
        #                 ['7870', 'KEPLER WEBER', 'KEPLER WEBER S.A.', ' '],
        #                 ['12653', 'KLABIN S/A', 'KLABIN S.A.', 'N2'],
        #                 ['24872', 'LIFEMED', 'LIFEMED INDUSTRIAL EQUIP. DE ART. MÉD. HOSP. S.A.', 'MA'],
        #                 ['19879', 'LIGHT S/A', 'LIGHT S.A.', 'NM'],
        #                 ['8036', 'LIGHT', 'LIGHT SERVICOS DE ELETRICIDADE S.A.', ' '],
        #                 ['23035', 'LINX', 'LINX S.A.', 'NM'], ['15091', 'LITEL', 'LITEL PARTICIPACOES S.A.', 'MB'],
        #                 ['24759', 'LITELA', 'LITELA PARTICIPAÇÕES S.A.', 'MB'],
        #                 ['19739', 'LOCALIZA', 'LOCALIZA RENT A CAR S.A.', 'NM'],
        #                 ['24910', 'LOCAWEB', 'LOCAWEB SERVIÇOS DE INTERNET S.A.', 'NM'],
        #                 ['23272', 'LOG COM PROP', 'LOG COMMERCIAL PROPERTIES', 'NM'],
        #                 ['20710', 'LOG-IN', 'LOG-IN LOGISTICA INTERMODAL S.A.', 'NM'],
        #                 ['8087', 'LOJAS AMERIC', 'LOJAS AMERICANAS S.A.', 'N1'],
        #                 ['8133', 'LOJAS RENNER', 'LOJAS RENNER S.A.', 'NM'], ['17434', 'LONGDIS', 'LONGDIS S.A.', 'MB'],
        #                 ['20370', 'LOPES BRASIL', 'LPS BRASIL - CONSULTORIA DE IMOVEIS S.A.', 'NM'],
        #                 ['20060', 'LUPATECH', 'LUPATECH S.A.', 'NM'],
        #                 ['20338', 'M.DIASBRANCO', 'M.DIAS BRANCO S.A. IND COM DE ALIMENTOS', 'NM'],
        #                 ['23612', 'MAESTROLOC', 'MAESTRO LOCADORA DE VEICULOS S.A.', 'MA'],
        #                 ['22470', 'MAGAZ LUIZA', 'MAGAZINE LUIZA S.A.', 'NM'],
        #                 ['8575', 'METAL LEVE', 'MAHLE-METAL LEVE S.A.', 'NM'],
        #                 ['8397', 'MANGELS INDL', 'MANGELS INDUSTRIAL S.A.', ' '],
        #                 ['8427', 'ESTRELA', 'MANUFATURA DE BRINQUEDOS ESTRELA S.A.', ' '],
        #                 ['8451', 'MARCOPOLO', 'MARCOPOLO S.A.', 'N2'],
        #                 ['20788', 'MARFRIG', 'MARFRIG GLOBAL FOODS S.A.', 'NM'],
        #                 ['22055', 'LOJAS MARISA', 'MARISA LOJAS S.A.', 'NM'],
        #                 ['8540', 'MERC FINANC', 'MERCANTIL BRASIL FINANC S.A. C.F.I.', ' '],
        #                 ['20613', 'METALFRIO', 'METALFRIO SOLUTIONS S.A.', 'NM'],
        #                 ['8605', 'METAL IGUACU', 'METALGRAFICA IGUACU S.A.', ' '],
        #                 ['8656', 'GERDAU MET', 'METALURGICA GERDAU S.A.', 'N1'],
        #                 ['13439', 'RIOSULENSE', 'METALURGICA RIOSULENSE S.A.', ' '],
        #                 ['8753', 'METISA', 'METISA METALURGICA TIMBOENSE S.A.', ' '],
        #                 ['22942', 'MGI PARTICIP', 'MGI - MINAS GERAIS PARTICIPAÇÕES S.A.', ' '],
        #                 ['22012', 'MILLS', 'MILLS ESTRUTURAS E SERVIÇOS DE ENGENHARIA S.A.', 'NM'],
        #                 ['8818', 'MINASMAQUINA', 'MINASMAQUINAS S.A.', ' '], ['20931', 'MINERVA', 'MINERVA S.A.', 'NM'],
        #                 ['13765', 'MINUPAR', 'MINUPAR PARTICIPACOES S.A.', ' '],
        #                 ['24902', 'MITRE REALTY', 'MITRE REALTY EMPREENDIMENTOS E PARTICIPAÇÕES S.A.', 'NM'],
        #                 ['17914', 'MMX MINER', 'MMX MINERACAO E METALICOS S.A.', 'NM'],
        #                 ['8893', 'MONT ARANHA', 'MONTEIRO ARANHA S.A.', ' '],
        #                 ['21067', 'MOURA DUBEUX', 'MOURA DUBEUX ENGENHARIA S/A', 'NM'],
        #                 ['23825', 'MOVIDA', 'MOVIDA PARTICIPACOES SA', 'NM'],
        #                 ['17949', 'MRS LOGIST', 'MRS LOGISTICA S.A.', 'MB'],
        #                 ['20915', 'MRV', 'MRV ENGENHARIA E PARTICIPACOES S.A.', 'NM'],
        #                 ['20982', 'MULTIPLAN', 'MULTIPLAN - EMPREEND IMOBILIARIOS S.A.', 'N2'],
        #                 ['5312', 'MUNDIAL', 'MUNDIAL S.A. - PRODUTOS DE CONSUMO', ' '],
        #                 ['24783', 'GRUPO NATURA', 'NATURA &CO HOLDING S.A.', 'NM'],
        #                 ['19550', 'NATURA', 'NATURA COSMETICOS S.A.', ' '],
        #                 ['15539', 'NEOENERGIA', 'NEOENERGIA S.A.', 'NM'],
        #                 ['9083', 'NORDON MET', 'NORDON INDUSTRIAS METALURGICAS S.A.', ' '],
        #                 ['22985', 'NORTCQUIMICA', 'NORTEC QUÍMICA S.A.', 'MA'],
        #                 ['24384', 'INTERMEDICA', 'NOTRE DAME INTERMEDICA PARTICIPACOES SA', 'NM'],
        #                 ['21334', 'NUTRIPLANT', 'NUTRIPLANT INDUSTRIA E COMERCIO S.A.', 'MA'],
        #                 ['22390', 'OCTANTE SEC', 'OCTANTE SECURITIZADORA S.A.', ' '],
        #                 ['20125', 'ODONTOPREV', 'ODONTOPREV S.A.', 'NM'], ['11312', 'OI', 'OI S.A.', 'N1'],
        #                 ['23426', 'OMEGA GER', 'OMEGA GERAÇÃO S.A.', 'NM'],
        #                 ['16942', 'OPPORT ENERG', 'OPPORTUNITY ENERGIA E PARTICIPACOES S.A.', 'MB'],
        #                 ['21342', 'OSX BRASIL', 'OSX BRASIL S.A.', 'NM'],
        #                 ['22250', 'OURINVESTSEC', 'OURINVEST SECURITIZADORA SA', ' '],
        #                 ['23507', 'OUROFINO S/A', 'OURO FINO SAUDE ANIMAL PARTICIPACOES S.A.', 'NM'],
        #                 ['23280', 'OURO VERDE', 'OURO VERDE LOCACAO E SERVICO S.A.', ' '],
        #                 ['94', 'PANATLANTICA', 'PANATLANTICA S.A.', ' '], ['20729', 'PARANA', 'PARANA BCO S.A.', ' '],
        #                 ['9393', 'PARANAPANEMA', 'PARANAPANEMA S.A.', 'NM'],
        #                 ['18236', 'PATRIA SEC', 'PATRIA CIA SECURITIZADORA DE CRED IMOB', ' '],
        #                 ['13773', 'PORTOBELLO', 'PBG S/A', 'NM'],
        #                 ['21644', 'PDG SECURIT', 'PDG COMPANHIA SECURITIZADORA', ' '],
        #                 ['20478', 'PDG REALT', 'PDG REALTY S.A. EMPREEND E PARTICIPACOES', 'NM'],
        #                 ['22187', 'PETRORIO', 'PETRO RIO S.A.', 'NM'],
        #                 ['24295', 'PETROBRAS BR', 'PETROBRAS DISTRIBUIDORA S/A', 'NM'],
        #                 ['9512', 'PETROBRAS', 'PETROLEO BRASILEIRO S.A. PETROBRAS', 'N2'],
        #                 ['9539', 'PETTENATI', 'PETTENATI S.A. INDUSTRIA TEXTIL', ' '],
        #                 ['13471', 'PLASCAR PART', 'PLASCAR PARTICIPACOES INDUSTRIAIS S.A.', ' '],
        #                 ['22160', 'POLO CAP SEC', 'POLO CAPITAL SECURITIZADORA S.A', ' '],
        #                 ['13447', 'POLPAR', 'POLPAR S.A.', ' '], ['19658', 'POMIFRUTAS', 'POMIFRUTAS S/A', 'NM'],
        #                 ['16659', 'PORTO SEGURO', 'PORTO SEGURO S.A.', 'NM'],
        #                 ['23523', 'PORTO VM', 'PORTO SUDESTE V.M. S.A.', ' '],
        #                 ['20362', 'POSITIVO TEC', 'POSITIVO TECNOLOGIA S.A.', 'NM'],
        #                 ['80152', 'PPLA', 'PPLA PARTICIPATIONS LTD.', 'DR3'],
        #                 ['24546', 'PRATICA', 'PRATICA KLIMAQUIP INDUSTRIA E COMERCIO SA', 'M2'],
        #                 ['24236', 'PRINER', 'PRINER SERVIÇOS INDUSTRIAIS S.A.', 'NM'],
        #                 ['19232', 'PROMAN', 'PRODUTORES ENERGET.DE MANSO S.A.- PROMAN', 'MB'],
        #                 ['20346', 'PROFARMA', 'PROFARMA DISTRIB PROD FARMACEUTICOS S.A.', 'NM'],
        #                 ['18333', 'PROMPT PART', 'PROMPT PARTICIPACOES S.A.', 'MB'],
        #                 ['22497', 'QUALICORP', 'QUALICORP CONSULTORIA E CORRETORA DE SEGUROS S.A.', 'NM'],
        #                 ['23302', 'QUALITY SOFT', 'QUALITY SOFTWARE S.A.', 'MA'],
        #                 ['5258', 'RAIADROGASIL', 'RAIA DROGASIL S.A.', 'NM'],
        #                 ['23230', 'RAIZEN ENERG', 'RAIZEN ENERGIA S.A.', ' '],
        #                 ['14109', 'RANDON PART', 'RANDON S.A. IMPLEMENTOS E PARTICIPACOES', 'N1'],
        #                 ['18406', 'RBCAPITALRES', 'RB CAPITAL COMPANHIA DE SECURITIZAÇÃO', 'MB'],
        #                 ['18430', 'WTORRE PIC', 'REAL AI PIC SEC DE CREDITOS IMOBILIARIO S.A.', ' '],
        #                 ['12572', 'RECRUSUL', 'RECRUSUL S.A.', ' '],
        #                 ['3190', 'REDE ENERGIA', 'REDE ENERGIA PARTICIPAÇÕES S.A.', ' '],
        #                 ['9989', 'PET MANGUINH', 'REFINARIA DE PETROLEOS MANGUINHOS S.A.', ' '],
        #                 ['21636', 'RENOVA', 'RENOVA ENERGIA S.A.', 'N2'],
        #                 ['21440', 'LE LIS BLANC', 'RESTOQUE COMÉRCIO E CONFECÇÕES DE ROUPAS S.A.', 'NM'],
        #                 ['16527', 'AES SUL', 'RGE SUL DISTRIBUIDORA DE ENERGIA S.A.', ' '],
        #                 ['18368', 'GER PARANAP', 'RIO PARANAPANEMA ENERGIA S.A.', ' '],
        #                 ['20451', 'RNI', 'RNI NEGÓCIOS IMOBILIÁRIOS S.A.', 'NM'],
        #                 ['23167', 'ROD COLINAS', 'RODOVIAS DAS COLINAS S.A.', ' '],
        #                 ['16306', 'ROSSI RESID', 'ROSSI RESIDENCIAL S.A.', 'NM'],
        #                 ['15300', 'ALL NORTE', 'RUMO MALHA NORTE S.A.', 'MB'],
        #                 ['17930', 'ALL PAULISTA', 'RUMO MALHA PAULISTA S.A.', 'MB'],
        #                 ['17450', 'RUMO S.A.', 'RUMO S.A.', 'NM'],
        #                 ['23540', 'SALUS INFRA', 'SALUS INFRAESTRUTURA PORTUARIA SA', ' '],
        #                 ['19593', 'SANESALTO', 'SANESALTO SANEAMENTO S.A.', ' '],
        #                 ['12696', 'SANSUY', 'SANSUY S.A. INDUSTRIA DE PLASTICOS', ' '],
        #                 ['14923', 'SANTHER', 'SANTHER FAB DE PAPEL STA THEREZINHA S.A.', ' '],
        #                 ['23388', 'STO ANTONIO', 'SANTO ANTONIO ENERGIA S.A.', ' '],
        #                 ['17892', 'SANTOS BRP', 'SANTOS BRASIL PARTICIPACOES S.A.', 'NM'],
        #                 ['13781', 'SAO CARLOS', 'SAO CARLOS EMPREEND E PARTICIPACOES S.A.', 'NM'],
        #                 ['20516', 'SAO MARTINHO', 'SAO MARTINHO S.A.', 'NM'],
        #                 ['9415', 'SPTURIS', 'SAO PAULO TURISMO S.A.', ' '],
        #                 ['10472', 'SARAIVA LIVR', 'SARAIVA LIVREIROS S.A. - EM RECUPERAÇÃO JUDICIAL', 'N2'],
        #                 ['14664', 'SCHULZ', 'SCHULZ S.A.', ' '], ['23221', 'SER EDUCA', 'SER EDUCACIONAL S.A.', 'NM'],
        #                 ['12823', 'ALIPERTI', 'SIDERURGICA J. L. ALIPERTI S.A.', ' '],
        #                 ['22799', 'SINQIA', 'SINQIA S.A.', 'NM'], ['20745', 'SLC AGRICOLA', 'SLC AGRICOLA S.A.', 'NM'],
        #                 ['24260', 'SMART FIT', 'SMARTFIT ESCOLA DE GINÁSTICA E DANÇA S.A.', 'M2'],
        #                 ['24252', 'SMILES', 'SMILES FIDELIDADE S.A.', 'NM'],
        #                 ['10880', 'SONDOTECNICA', 'SONDOTECNICA ENGENHARIA SOLOS S.A.', ' '],
        #                 ['10960', 'SPRINGER', 'SPRINGER S.A.', ' '],
        #                 ['20966', 'SPRINGS', 'SPRINGS GLOBAL PARTICIPACOES S.A.', 'NM'],
        #                 ['24201', 'STARA', 'STARA S.A. - INDÚSTRIA DE IMPLEMENTOS AGRÍCOLAS', 'MA'],
        #                 ['22594', 'STATKRAFT', 'STATKRAFT ENERGIAS RENOVAVEIS S.A.', ' '],
        #                 ['16586', 'SUDESTE S/A', 'SUDESTE S.A.', 'MB'],
        #                 ['16438', 'SUL 116 PART', 'SUL 116 PARTICIPACOES S.A.', 'MB'],
        #                 ['21121', 'SUL AMERICA', 'SUL AMERICA S.A.', 'N2'],
        #                 ['9067', 'SUZANO HOLD', 'SUZANO HOLDING S.A.', ' '],
        #                 ['13986', 'SUZANO S.A.', 'SUZANO S.A.', 'NM'],
        #                 ['22454', 'TIME FOR FUN', 'T4F ENTRETENIMENTO S.A.', 'NM'],
        #                 ['6173', 'TAURUS ARMAS', 'TAURUS ARMAS S.A.', 'N2'],
        #                 ['24066', 'TCP TERMINAL', 'TCP TERMINAL DE CONTEINERES DE PARANAGUA SA', ' '],
        #                 ['22519', 'TECHNOS', 'TECHNOS S.A.', 'NM'], ['20435', 'TECNISA', 'TECNISA S.A.', 'NM'],
        #                 ['11207', 'TECNOSOLO', 'TECNOSOLO ENGENHARIA S.A.', ' '],
        #                 ['20800', 'TEGMA', 'TEGMA GESTAO LOGISTICA S.A.', 'NM'],
        #                 ['11223', 'TEKA', 'TEKA-TECELAGEM KUEHNRICH S.A.', ' '],
        #                 ['11231', 'TEKNO', 'TEKNO S.A. - INDUSTRIA E COMERCIO', ' '],
        #                 ['11258', 'TELEBRAS', 'TELEC BRASILEIRAS S.A. TELEBRAS', ' '],
        #                 ['17671', 'TELEF BRASIL', 'TELEFÔNICA BRASIL S.A', ' '],
        #                 ['23329', 'TERM. PE III', 'TERMELÉTRICA PERNAMBUCO III S.A.', ' '],
        #                 ['18538', 'MENEZES CORT', 'TERMINAL GARAGEM MENEZES CORTES S.A.', 'MB'],
        #                 ['19852', 'TERMOPE', 'TERMOPERNAMBUCO S.A.', ' '],
        #                 ['20354', 'TERRA SANTA', 'TERRA SANTA AGRO S.A.', 'NM'],
        #                 ['7544', 'TEX RENAUX', 'TEXTIL RENAUXVIEW S.A.', ' '],
        #                 ['17639', 'TIM PART S/A', 'TIM PARTICIPACOES S.A.', 'NM'],
        #                 ['19992', 'TOTVS', 'TOTVS S.A.', 'NM'],
        #                 ['19330', 'TRIUNFO PART', 'TPI - TRIUNFO PARTICIP. E INVEST. S.A.', 'NM'],
        #                 ['20257', 'TAESA', 'TRANSMISSORA ALIANÇA DE ENERGIA ELÉTRICA S.A.', 'N2'],
        #                 ['8192', 'TREVISA', 'TREVISA INVESTIMENTOS S.A.', ' '],
        #                 ['23060', 'TRIANGULOSOL', 'TRIÂNGULO DO SOL AUTO-ESTRADAS S.A.', ' '],
        #                 ['21130', 'TRISUL', 'TRISUL S.A.', 'NM'],
        #                 ['11398', 'CRISTAL', 'TRONOX PIGMENTOS DO BRASIL S.A.', ' '],
        #                 ['22276', 'TRUESEC', 'TRUE SECURITIZADORA S.A.', ' '], ['6343', 'TUPY', 'TUPY S.A.', 'NM'],
        #                 ['18465', 'ULTRAPAR', 'ULTRAPAR PARTICIPACOES S.A.', 'NM'],
        #                 ['22780', 'UNICASA', 'UNICASA INDÚSTRIA DE MÓVEIS S.A.', 'NM'],
        #                 ['21555', 'UNIDAS', 'UNIDAS S.A.', ' '], ['11592', 'UNIPAR', 'UNIPAR CARBOCLORO S.A.', ' '],
        #                 ['16624', 'UPTICK', 'UPTICK PARTICIPACOES S.A.', 'MB'],
        #                 ['14320', 'USIMINAS', 'USINAS SID DE MINAS GERAIS S.A.-USIMINAS', 'N1'],
        #                 ['4170', 'VALE', 'VALE S.A.', 'NM'], ['20028', 'VALID', 'VALID SOLUÇÕES S.A.', 'NM'],
        #                 ['23990', 'VERTCIASEC', 'VERT COMPANHIA SECURITIZADORA', ' '],
        #                 ['6505', 'VIAVAREJO', 'VIA VAREJO S.A.', 'NM'],
        #                 ['24805', 'VIVARA S.A.', 'VIVARA PARTICIPAÇOES S.A', 'NM'],
        #                 ['20702', 'VIVER', 'VIVER INCORPORADORA E CONSTRUTORA S.A.', 'NM'],
        #                 ['11762', 'VULCABRAS', 'VULCABRAS/AZALEIA S.A.', 'NM'], ['5410', 'WEG', 'WEG S.A.', 'NM'],
        #                 ['11991', 'WETZEL S/A', 'WETZEL S.A.', ' '], ['14346', 'WHIRLPOOL', 'WHIRLPOOL S.A.', ' '],
        #                 ['80047', 'WILSON SONS', 'WILSON SONS LTD.', 'DR3'],
        #                 ['23590', 'WIZ S.A.', 'WIZ SOLUÇÕES E CORRETAGEM DE SEGUROS S.A.', 'NM'],
        #                 ['11070', 'WLM IND COM', 'WLM PART. E COMÉRCIO DE MÁQUINAS E VEÍCULOS S.A.', ' '],
        #                 ['21016', 'YDUQS PART', 'YDUQS PARTICIPACOES S.A.', 'NM']]
        # print('debug b3_companies', b3_companies)
        # return b3_companies

        print('LOAD page of companies')
        url = main_url
        btnTodas = '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_btnTodas"]'
        table_company = '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]'
        row_company = '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr'
        col_company = '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[1]/td'

        browser.get(url)
        browser.minimize_window()

        # load
        assert (EC.element_to_be_clickable((By.XPATH, btnTodas)))
        wait.until(EC.element_to_be_clickable((By.XPATH, btnTodas)))
        # click
        browser.find_element(By.XPATH, btnTodas).click()
        # load
        assert (EC.presence_of_element_located((By.XPATH, table_company)))
        wait.until(EC.presence_of_element_located((By.XPATH, table_company)))
        # get lenght
        rowsA = len(browser.find_elements(By.XPATH, row_company))
        colsA = len(browser.find_elements(By.XPATH, col_company))
        # get data
        for r in range(1, rowsA + 1):
            row = [browser.find_element(By.XPATH,
                                        '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[' + str(
                                            r) + ']/td[1]/a').get_attribute('href').split("=") [1],
                   browser.find_element_by_xpath(
                       '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[' + str(
                           r) + ']/td[2]').text, browser.find_element_by_xpath(
                    '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[' + str(
                        r) + ']/td[1]').text, browser.find_element_by_xpath(
                    '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[' + str(
                        r) + ']/td[3]').text]
            b3_companies.append(row)
            if r % 50 == 0:
                print(r)
        print(rowsA)
        print('...done')
        return b3_companies
    except Exception as e:
        restart(e, __name__)
def getCompanyMainPage(company):
    try:
        print('...get details', company [col ['PREGÃO']])
        browser.get(cvm_url + company [col ['CMV']])
        # browser.minimize_window()

        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body')))
        company [col [
            'LINK']] = 'http://bvmf.bmfbovespa.com.br/cias-listadas/empresas-listadas/ResumoEmpresaPrincipal.aspx?codigoCvm=' + \
                       company [col ['CMV']]
        company [col ['DATA']] = datetime.strptime(
            getCompanyItem('/html/body/', 'div[1]/div[1]').replace('Atualizado em ', '').replace(', às', ''),
            '%d/%m/%Y %Hh%M').strftime("%d/%m/%Y %H:%M:00")
        company [col ['SITE']] = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[6]/td[2]')
        company [col ['CNPJ']] = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[3]/td[2]')
        company [col ['TICKER']] = ' '.join(re.split(' ', getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div',
                                                                         '/table/tbody/tr[2]/td[2]').replace(
            'Mais Códigos\n', '').replace(';', '')))
        setores = re.split(' / ', getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[5]/td[2]'))
        company [col ['SETOR']] = setores [0]
        company [col ['SUBSETOR']] = setores [1]
        company [col ['SEGMENTO']] = setores [2]
        company [col ['ATIVIDADE']] = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div',
                                                     '/table/tbody/tr[4]/td[2]')
        company [col ['LOG']] = '1 company'
        company [col ['TIMESTAMP']] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        return company
    except Exception as e:
        restart(e, __name__)
def getCompanyItem(base_xpath, xpath):
    try:
        try:
            item = browser.find_element(By.XPATH, base_xpath + xpath).text
        except:
            item = browser.find_element(By.XPATH, base_xpath + '/div/div[1]' + xpath).text
        return item
    except Exception as e:
        pass
def createSheetReports(company):
    try:
        name = {'name': 'R - ' + str(
            company [col ['SETOR']] [:5] + ' / ' + company [col ['SUBSETOR']] [:5] + ' / ' + company [
                                                                                                 col ['SEGMENTO']] [
                                                                                             :5] + ' - ' + company [
                                                                                                               col [
                                                                                                                   'TICKER']] [
                                                                                                           :4].replace(
                'Nenh',
                'NONE') + ' - ' +
            company [col ['PREGÃO']] + ' - ' + company [col ['EMPRESA']] + ' - ' + company [col ['CMV']])}
        newsheet = gdrive.files().copy(fileId=report_sheet, body=name).execute()
        company_id_reports = newsheet.get('id')
        company [col ['REPORTS']] = sheet_url + company_id_reports
        permission_user = {
            'type': 'user',
            'role': 'writer',
            'emailAddress': CLIENT_EMAIL
        }
        print('PLEASE REMEMBER TO REACTIVATE PERMISSIONS IN 24h')
        # permission = gdrive.permissions().create(fileId=company_id_reports, body=permission_user, fields="id").execute()
        # permission_id = permission.get('id')

        print('...create sheet', company [col ['PREGÃO']], company_id_reports)
        return company
    except Exception as e:
        restart(e, __name__)
def createSheetFundamentos(company):
    try:
        name = {'name': 'A - ' + str(
            company [col ['SETOR']] [:5] + ' / ' + company [col ['SUBSETOR']] [:5] + ' / ' + company [
                                                                                                 col ['SEGMENTO']] [
                                                                                             :5] + ' - ' + company [
                                                                                                               col [
                                                                                                                   'TICKER']] [
                                                                                                           :4].replace(
                'Nenh',
                'NONE') + ' - ' +
            company [col ['PREGÃO']] + ' - ' + company [col ['EMPRESA']] + ' - ' + company [col ['CMV']])}
        newsheet = gdrive.files().copy(fileId=fundament_sheet, body=name).execute()
        company_id_fundamentos = newsheet.get('id')
        company [col ['FUNDAMENTOS']] = sheet_url + company_id_fundamentos
        permission_user = {
            'type': 'user',
            'role': 'writer',
            'emailAddress': CLIENT_EMAIL
        }
        print('PLEASE REMEMBER TO REACTIVATE PERMISSIONS IN 24h')
        # permission = gdrive.permissions().create(fileId=company_id_reports, body=permission_user, fields="id").execute()
        # permission_id = permission.get('id')

        print('...create sheet', company [col ['PREGÃO']], company_id_fundamentos)
        return company
    except Exception as e:
        restart(e, __name__)
def setCompanyMainPage(company):
    try:
        global google
        if google != True:
            google = googleAPI()
        loadCompanyReportsSheets(company)

        # update company sheet
        report_sheet_index.resize(6, 6)
        report_sheet_index.update_acell('F4', company [col ['CMV']])
        report_sheet_index.update_acell('A1', company [col ['PREGÃO']])
        report_sheet_index.update_acell('F1', company [col ['EMPRESA']])
        report_sheet_index.update_acell('D4', company [col ['MERCADO']])
        # report_sheet_index.update_acell('A1', company[col['LINK']])
        # report_sheet_index.update_acell('A1', company[col['REPORTS']])
        report_sheet_index.update_acell('F2', company [col ['DATA']])
        report_sheet_index.update_acell('A3', company [col ['SITE']])
        report_sheet_index.update_acell('F3', company [col ['CNPJ']])
        report_sheet_index.update_acell('E4', company [col ['TICKER']])
        report_sheet_index.update_acell('A4', company [col ['SETOR']])
        report_sheet_index.update_acell('B4', company [col ['SUBSETOR']])
        report_sheet_index.update_acell('C4', company [col ['SEGMENTO']])
        report_sheet_index.update_acell('A2', company [col ['ATIVIDADE']])
        print('...set details', company [col ['PREGÃO']])
        return company
    except Exception as e:
        restart(e, __name__)
def setCompanyFundamentos(company):
    try:
        global google
        if google != True:
            google = googleAPI()
        loadCompanyFundamentosSheets(company)
        # global fundamentos_sheet_index
        # global fundamentos_sheet_reports

        # update company sheet
        fundamentos_sheet_index.resize(6, 6)
        fundamentos_sheet_reports.resize(2, 9)
        range1 = 'index!A:ZZZ'
        range2 = 'reports!A2:H'
        fundamentos_sheet_index.update_acell('A1', '=IMPORTRANGE("' + company [col ['REPORTS']] + '";"' + range1 + '")')
        fundamentos_sheet_reports.update_acell('A2', '=IMPORTRANGE("' + company [col ['REPORTS']] + '";"' + range2 + '")')

        print('...set details', company [col ['PREGÃO']])
        return company
    except Exception as e:
        restart(e, __name__)
def setb3Company(company):
    try:
        global google
        if google != True:
            google = googleAPI()

        print('...update listagem', company [col ['PREGÃO']])
        return company
    except Exception as e:
        restart(e, __name__)
def getSheetCompanyReports(company):
    try:
        loadCompanyReportsSheets(company)

        sheet_report_list = report_sheet_index.get_all_values()
        # sheet_data = report_sheet_reports.get_all_values()
        sheet_report_list = [line for line in sheet_report_list if 'bmfbovespa.com.br' in line [0] or 'rad.cvm.gov.br' in line [0]]
        for r, row in enumerate(sheet_report_list):
            del sheet_report_list [r] [-1]
        return sheet_report_list
    except Exception as e:
        restart(e, __name__)
def getb3CompanyReports(company):
    try:
        b3_reports = []
        dfp = getb3CompanyReportList(company, 'dfp')
        # dfp = [[
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=81946&CodigoTipoInstituicao=2',
        #     '81946', '31/12/2018', 'Demonstrações Financeiras Padronizadas', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=72981&CodigoTipoInstituicao=2',
        #     '72981', '31/12/2017', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=64150&CodigoTipoInstituicao=2',
        #     '64150', '31/12/2016', 'Demonstrações Financeiras Padronizadas', 'Versão 3.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=54919&CodigoTipoInstituicao=2',
        #     '54919', '31/12/2015', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=47330&CodigoTipoInstituicao=2',
        #     '47330', '31/12/2014', 'Demonstrações Financeiras Padronizadas', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=35867&CodigoTipoInstituicao=2',
        #     '35867', '31/12/2013', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=26297&CodigoTipoInstituicao=2',
        #     '26297', '31/12/2012', 'Demonstrações Financeiras Padronizadas', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=15636&CodigoTipoInstituicao=2',
        #     '15636', '31/12/2011', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=5715&CodigoTipoInstituicao=2',
        #     '5715', '31/12/2010', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=31/12/2009&tipo=2',
        #     '21725', '31/12/2009', 'Demonstrações Financeiras Padronizadas', 'Apresentação'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=31/12/2008&tipo=2',
        #     '21725', '31/12/2008', 'Demonstrações Financeiras Padronizadas', 'Apresentação']]
        # print('debug dfp', dfp)
        b3_reports.extend(dfp)
        itr = getb3CompanyReportList(company, 'itr')
        # itr = [[
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=89680&CodigoTipoInstituicao=2',
        #     '89680', '30/09/2019', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=86702&CodigoTipoInstituicao=2',
        #     '86702', '30/06/2019', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=83730&CodigoTipoInstituicao=2',
        #     '83730', '31/03/2019', 'Informações Trimestrais', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=80252&CodigoTipoInstituicao=2',
        #     '80252', '30/09/2018', 'Informações Trimestrais', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=78058&CodigoTipoInstituicao=2',
        #     '78058', '30/06/2018', 'Informações Trimestrais', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=74580&CodigoTipoInstituicao=2',
        #     '74580', '31/03/2018', 'Informações Trimestrais', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=71504&CodigoTipoInstituicao=2',
        #     '71504', '30/09/2017', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=70223&CodigoTipoInstituicao=2',
        #     '70223', '30/06/2017', 'Informações Trimestrais', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=66982&CodigoTipoInstituicao=2',
        #     '66982', '31/03/2017', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=61030&CodigoTipoInstituicao=2',
        #     '61030', '30/09/2016', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=59081&CodigoTipoInstituicao=2',
        #     '59081', '30/06/2016', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=56612&CodigoTipoInstituicao=2',
        #     '56612', '31/03/2016', 'Informações Trimestrais', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=51881&CodigoTipoInstituicao=2',
        #     '51881', '30/09/2015', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=50475&CodigoTipoInstituicao=2',
        #     '50475', '30/06/2015', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=48431&CodigoTipoInstituicao=2',
        #     '48431', '31/03/2015', 'Informações Trimestrais', 'Versão 2.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=42676&CodigoTipoInstituicao=2',
        #     '42676', '30/09/2014', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=40652&CodigoTipoInstituicao=2',
        #     '40652', '30/06/2014', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=37603&CodigoTipoInstituicao=2',
        #     '37603', '31/03/2014', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=33117&CodigoTipoInstituicao=2',
        #     '33117', '30/09/2013', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=30782&CodigoTipoInstituicao=2',
        #     '30782', '30/06/2013', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=27303&CodigoTipoInstituicao=2',
        #     '27303', '31/03/2013', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=22651&CodigoTipoInstituicao=2',
        #     '22651', '30/09/2012', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=20619&CodigoTipoInstituicao=2',
        #     '20619', '30/06/2012', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=19334&CodigoTipoInstituicao=2',
        #     '19334', '31/03/2012', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=13049&CodigoTipoInstituicao=2',
        #     '13049', '30/09/2011', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=11121&CodigoTipoInstituicao=2',
        #     '11121', '30/06/2011', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=8089&CodigoTipoInstituicao=2',
        #     '8089', '31/03/2011', 'Informações Trimestrais', 'Versão 1.0'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=30/09/2010&tipo=4',
        #     '21725', '30/09/2010', 'Informações Trimestrais', 'Reapresentação'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=30/06/2010&tipo=4',
        #     '21725', '30/06/2010', 'Informações Trimestrais', 'Reapresentação'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=31/03/2010&tipo=4',
        #     '21725', '31/03/2010', 'Informações Trimestrais', 'Reapresentação'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=30/09/2009&tipo=4',
        #     '21725', '30/09/2009', 'Informações Trimestrais', 'Apresentação'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=30/06/2009&tipo=4',
        #     '21725', '30/06/2009', 'Informações Trimestrais', 'Apresentação'], [
        #     'http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=1&razao=ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.&pregao=ADVANCED-DH&ccvm=21725&data=31/03/2009&tipo=4',
        #     '21725', '31/03/2009', 'Informações Trimestrais', 'Apresentação']]
        # print('debug itr', itr)
        b3_reports.extend(itr)

        for i, item in enumerate(b3_reports):
            data = datetime.strptime(item [2], '%d/%m/%Y')
            b3_reports [i] [2] = data.strftime('%Y') + '/' + data.strftime('%m') + '/' + data.strftime('%d')
        return b3_reports
    except Exception as e:
        restart(e, __name__)
def getb3CompanyReportList(company, report):
    try:
        print('... Company Reports', report)
        link = 'http://bvmf.bmfbovespa.com.br/cias-listadas/empresas-listadas/HistoricoFormularioReferencia.aspx?codigoCVM=' + \
               company [col ['CMV']] + '&tipo=' + report + '&ano=0&idioma=pt-br'
        # load company report page
        browser.get(link)
        browser.minimize_window()

        assert (
            EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_contentPlaceHolderConteudo_divDemonstrativo"]')))
        wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_contentPlaceHolderConteudo_divDemonstrativo"]')))
        links = browser.find_elements(By.TAG_NAME, 'a')
        list_of_reports = reportLinkParser(company, links)
        return list_of_reports
    except Exception as e:
        restart(e, __name__)
def reportLinkParser(company, links):
    try:
        list_of_reports = []
        pub_date = ''
        for link in links:
            # check if report is newer
            text = str(link.text)
            href = str(link.get_attribute('href'))
            if text:
                if pub_date != text.split(" - ") [0]:
                    # parse new style links
                    if 'AbreFormularioCadastral' in href:
                        report = href.replace("&CodigoTipoInstituicao=2')", "").replace(
                            "javascript:AbreFormularioCadastral('http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=",
                            "")
                        link = href.replace("')", "").replace("javascript:AbreFormularioCadastral('", "")
                        report_item = [link, report, text.split(" - ") [0], text.split(" - ") [1],
                                       text.split(" - ") [2]]
                        list_of_reports.append(report_item)
                    # parse old style links
                    if 'ConsultarDXW' in href:
                        report = href.replace("')", "").replace(
                            "javascript:ConsultarDXW('http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?", "")
                        link = href.replace("')", "").replace("javascript:ConsultarDXW('", "")
                        report_item = [link, company [col ['CMV']], text.split(" - ") [0], text.split(" - ") [1],
                                       text.split(" - ") [2]]
                        list_of_reports.append(report_item)
                pub_date = text.split(" - ") [0]
        return list_of_reports
    except Exception as e:
        restart(e, __name__)
def reportContentRAD_Dados(company, row):
    # this is where the magic gets downloaded
    # company
    # url = row[0]
    # relat = row[2]
    # optA = row[3]
    # optB = row[4]
    global url
    global google
    if google != True:
        google = googleAPI()
    try:

        if url != row[0]:
            browser.get(row[0])
        if 'Dados' in row[2]:
            xpathA = '//*[@id="cmbGrupo"]'
            xpathB = '//*[@id="cmbQuadro"]'
            frame_name = 'iFrameFormulariosFilho'
            frame = '//*[@id="iFrameFormulariosFilho"]'
            table = '//*[@id="UltimaTabela"]'

        browser.switch_to.default_content()
        # find correct options A
        browser.find_element(By.XPATH, xpathA)
        wait.until(EC.presence_of_element_located((By.XPATH, xpathA)))
        selectA = Select(browser.find_element(By.XPATH, xpathA))
        selectA.select_by_visible_text(row [3])
        time.sleep(2)

        # find correct options B
        browser.find_element(By.XPATH, xpathB)
        wait.until(EC.presence_of_element_located((By.XPATH, xpathB)))
        selectB = Select(browser.find_element(By.XPATH, xpathB))
        selectB.select_by_visible_text(row [4])
        time.sleep(2)

        # focus and prepare for content
        wait.until(EC.presence_of_element_located((By.XPATH, frame)))
        browser.switch_to.frame(frame_name)
        time.sleep(2)

        wait.until(EC.presence_of_element_located((By.XPATH, table)))
        time.sleep(2)

        report = []
        col = []
        col.append(company[1].strip())
        col.append(row[1].strip())
        col.append(row[2].strip())
        col.append(row[3].strip())
        col.append(row[4].strip())
        col.append('Ação')
        col.append('ON')
        col3 = browser.find_element_by_xpath('//*[@id="QtdAordCapiItgz_1"]').text
        col.append(col3.strip())
        report.append(col)

        col = []
        col.append(company[1].strip())
        col.append(row[1].strip())
        col.append(row[2].strip())
        col.append(row[3].strip())
        col.append(row[4].strip())
        col.append('Ação')
        col.append('PN')
        col3 = browser.find_element_by_xpath('//*[@id="QtdAprfCapiItgz_1"]').text
        col.append(col3.strip())
        report.append(col)

        return report
    except Exception as e:
        restart(e, __name__)
def reportContentRAD_DRE(company, row):
    # this is where the magic gets downloaded
    # company
    # url = row[0]
    # relat = row[2]
    # optA = row[3]
    # optB = row[4]
    global google
    if google != True:
        google = googleAPI()

    global url
    global DRE
    DRE = ''
    try:

        if url != row[0]:
            browser.get(row[0])
        if 'DRE' in row[2]:
            xpathA = '//*[@id="cmbGrupo"]'
            xpathB = '//*[@id="cmbQuadro"]'
            frame_name = 'iFrameFormulariosFilho'
            frame = '//*[@id="iFrameFormulariosFilho"]'
            table = '//*[@id="ctl00_cphPopUp_tbDados"]'

        browser.switch_to.default_content()
        # find correct options A
        browser.find_element(By.XPATH, xpathA)
        wait.until(EC.presence_of_element_located((By.XPATH, xpathA)))
        selectA = Select(browser.find_element(By.XPATH, xpathA))
        selectA.select_by_visible_text(row[3])
        time.sleep(2)

        # find correct options B
        browser.find_element(By.XPATH, xpathB)
        wait.until(EC.presence_of_element_located((By.XPATH, xpathB)))
        selectB = Select(browser.find_element(By.XPATH, xpathB))
        selectB.select_by_visible_text(row[4])
        time.sleep(2)

        # focus and prepare for content
        wait.until(EC.presence_of_element_located((By.XPATH, frame)))
        browser.switch_to.frame(frame_name)
        time.sleep(2)

        wait.until(EC.presence_of_element_located((By.XPATH, table)))
        time.sleep(2)
        unidade = browser.find_element_by_xpath('//*[@id="TituloTabelaSemBorda"]').text
        if 'Reais Mil' in unidade:
            unidade = 1
        else:
            unidade = 1000
        rows: int = len(browser.find_elements(By.XPATH, '//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr'))
        cols: int = len(browser.find_elements(By.XPATH, '//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[1]/td'))

        txt = 'txt'
        if row[4] in ['Balanço Patrimonial Ativo', 'Balanço Patrimonial Passivo']:
            txt = "cell"

        # DATA GRABBING, FOR REAL
        report = []
        for r in range(2, rows + 1):
            col = []
            col.append(company[1].strip())
            col.append(row[1].strip())
            col.append(row[2].strip())
            col.append(row[3].strip())
            col.append(row[4].strip())
            col1 = browser.find_element_by_xpath('//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[' + str(r) + ']/td[1]').text
            col2 = browser.find_element_by_xpath('//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[' + str(r) + ']/td[2]').text
            col3 = browser.find_element_by_xpath('//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[' + str(r) + ']/td[3]').text
            if '\n a \n' in col3:
                col3 = col3.split('\n')[2] + ' -- ' + col3
            col.append(col1.strip())
            col.append(col2.strip())
            try:
                col.append(float(int(col3.replace('.','')) / int(unidade)))
            except:
                col.append(col3.strip())
            report.append(col)


        url = row[0]
        return report
    except Exception as e:
        DRE = row[3]
        return ''
def updateReportToSheet(company, report):
    try:
        global google
        if google != True:
            google = googleAPI()

        loadCompanyReportsSheets(company)

        # combine sheet_data and reports
        sheet_data = report_sheet_reports.get_all_values()
        sheet_data.extend(report)
        # deduplicate data
        sheet_data = dedupReport(sheet_data)
        # batch update to sheet
        data = sheet_data
        start_col = 1
        start_row = 1
        sheet = report_sheet_reports
        sheet_range = sheetRange(start_row, start_col, data)

        cell_list = sheet.range(sheet_range)
        for cell in cell_list:
            cell.value = data[cell.row-1][cell.col-1]
        sheet.resize(len(sheet_data), len(sheet_data[0])+1)
        sheet.update_cells(cell_list)
        logCompany(company, '3 data')
    except Exception as e:
        restart(e, __name__)
def dedupReport(report):
    try:
        # dummy separate dup and non-dup
        dup = []
        ndup = []
        for i, item in enumerate(report):
            dup.append([])
            ndup.append([])
            dup[i].append(item[0])
            dup[i].append(item[1])
            dup[i].append(item[2])
            dup[i].append(item[3])
            dup[i].append(item[4])
            dup[i].append(item[5])
            dup[i].append(item[6])
            ndup[i].append(item[7])
        # dummy deduplicate/remove duplicates
        dup_dedup = []
        ndup_dedup = []
        for l, line in enumerate(dup):
            if line not in dup_dedup:
                dup_dedup.append(dup[l])
                ndup_dedup.append(ndup[l])
        # dummy combine again original list
        report = []
        for r, row in enumerate(dup_dedup):
            report.append([])
            report[r].extend(dup_dedup[r])
            report[r].extend(ndup_dedup[r])

        # dummy multiple sort
        report.sort(key=lambda x: (x[5]), reverse=False)
        report.sort(key=lambda x: (x[4]), reverse=False)
        report.sort(key=lambda x: (x[3]), reverse=True)
        report.sort(key=lambda x: (x[2]), reverse=False)
        report.sort(key=lambda x: (x[1]), reverse=True)

        return report
    except Exception as e:
        restart(e, __name__)
def logCompany(company, log):
    try:
        company [col ['LOG']] = log
        company [col ['TIMESTAMP']] = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        row = bovespa_listagem.find(str(company [col ['CMV']])).row
        sheet_range = 'A' + str(row) + ':' + sheetCol(len(company)) + str(row)
        cell_list = bovespa_listagem.range(sheet_range)
        for c, cell in enumerate(cell_list):
            try:
                cell.value = company [c]
            except:
                pass
        bovespa_listagem.update_cells(cell_list)
        # print ('debug LOG not upddated')

        bovespa_log.append_row([company [col ['CMV']], company [col ['LOG']], company [col ['TIMESTAMP']]])
        return company
    except Exception as e:
        restart(e, __name__)


# MAIN BLOCKS
def companyList():
    try:
        sheet_companies = sheetCompany()
        b3_companies = b3Company()

        if sheet_companies:
            while len(sheet_companies [0]) > len(b3_companies [0]):
                for item in sheet_companies:
                    item.pop()

        print('UPDATE sheet of companies')

        found = list_intersection(b3_companies, sheet_companies)
        not_found = list_remove_extra(b3_companies, sheet_companies)

        if not_found:
            # batch add list
            start_row = sheet_companies.__len__() + 1
            start_col = 1
            end_row = start_row + not_found.__len__()
            end_col = bovespa_listagem.col_count
            bovespa_listagem.resize(end_row, end_col)

            sheet_range1 = sheetCol(start_col) + str(start_row + 1)
            sheet_range2 = sheetCol(end_col) + str(end_row)
            sheet_range = sheet_range1 + ':' + sheet_range2
            cell_list = bovespa_listagem.range(sheet_range)
            try:
                for cell in cell_list:
                    if cell.col == end_col - 1:
                        cell.value = '1 company'
                    if cell.col == end_col:
                        cell.value = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    try:
                        cell.value = not_found [cell.row - start_row - 1] [cell.col - 1]
                    except:
                        pass
                    # print(cell.row - start_row, cell.col -1)
            except:
                pass
            bovespa_listagem.update_cells(cell_list)

        print(not_found.__len__(), 'Companies updated to sheet')
        print('...done')
    except Exception as e:
        restart(e, __name__)
def companyOrder():
    try:
        global google
        if google != True:
            google = googleAPI()
        loadSheets()

        global sheet_companies
        print('ORDER BY list of companies')
        sheet_companies = bovespa_listagem.get_all_values()

        # order by timestamp newer
        try:
            sheet_companies = sorted(sheet_companies, key=lambda x: (x [col ['LOG']], x [col ['TIMESTAMP']]),
                                     reverse=False)
        except:
            pass
        print('done...')
        return sheet_companies
    except Exception as e:
        restart(e, __name__)
def companySheet(company):
    try:
        global google
        if google != True:
            google = googleAPI()
        print('GET COMPANY DATA for', company [col ['PREGÃO']])
        action = ''

        global company_id_reports
        if company [col ['REPORTS']]:
            company_id_reports = company [col ['REPORTS']].replace(sheet_url, '')
            action1 = 'found'
        else:
            company = getCompanyMainPage(company)
            company_id_reports = createSheetReports(company) [col ['REPORTS']].replace(sheet_url, '')
            company = setCompanyMainPage(company)
            action1 = 'created'
        global company_id_fundamentos
        if company [col ['FUNDAMENTOS']]:
            company_id_fundamentos = company [col ['FUNDAMENTOS']].replace(sheet_url, '')
            action2 = 'found'
        else:
            company_id_fundamentos = createSheetFundamentos(company) [col ['FUNDAMENTOS']].replace(sheet_url, '')
            company = setCompanyFundamentos(company)
            action2 = 'created'

        if action1 == 'created':
            if action2 == 'created':
                action ='created'
            else:
                action ='created and found'
        else:
            if action2 == 'created':
                action == 'found and created'
            else:
                action = 'found'

        company = setb3Company(company)
        logCompany(company, '1 company')

        print('...done', action, company [col ['PREGÃO']], sheet_url + company [col ['REPORTS'], company [col ['FUNDAMENTOS'])
        return company
    except Exception as e:
        restart(e, __name__)
def companyReportList(company):
    try:
        global google
        if google != True:
            google = googleAPI()
        print('GET REPORTS for', company [col ['PREGÃO']])

        sheet_reports_list = getSheetCompanyReports(company)
        b3_reports_list = getb3CompanyReports(company)

        company_reports_list = list_unique(sheet_reports_list, b3_reports_list)
        company_reports_list = sorted(company_reports_list, key=lambda x: datetime.strptime(x [2], '%Y/%m/%d'),
                                      reverse=True)

        start_row = 6
        start_col = 1
        end_row = start_row + company_reports_list.__len__()
        end_col = report_sheet_index.col_count
        report_sheet_index.resize(end_row, end_col)

        sheet_range1 = sheetCol(start_col) + str(start_row)
        sheet_range2 = sheetCol(company_reports_list [0].__len__()) + str(end_row)
        sheet_range = sheet_range1 + ':' + sheet_range2

        cell_list = report_sheet_index.range(sheet_range)
        for cell in cell_list:
            try:
                r = cell.row - start_row
                c = cell.col - start_col
                value = company_reports_list [r] [c]
                cell.value = value
            except:
                pass
        report_sheet_index.update_cells(cell_list)

        logCompany(company, '2 report')

        print('...done')
        return company_reports_list
    except Exception as e:
        restart(e, __name__)
def companyReportListData(company, company_reports_list):
    try:
        global google
        if google != True:
            google = googleAPI()

        print('GET DATA REPORT for', company [col ['PREGÃO']])
        loadCompanyReportsSheets(company)

        sheet_data = report_sheet_reports.get_all_values()
        del sheet_data [0]
        sheet_reports_list = list_unique([], [[r [1], r [2], r [3], r [4]] for r in sheet_data])

        company_reports_list2 = [[c [2]] + p for p in parts for c in company_reports_list]
        company_reports_list2 = sorted(company_reports_list2, key=lambda x: (x [0]), reverse=True)

        missing_reports = list_difference(sheet_reports_list, company_reports_list2)
        existing_reports = list_intersection(sheet_reports_list, company_reports_list2)

        for i1, item1 in enumerate(existing_reports):
            for l1, line1 in enumerate(company_reports_list):
                if item1 [0] == line1 [2]:
                    existing_reports [i1].insert(0, line1 [0])
        for i2, item2 in enumerate(missing_reports):
            for l2, line2 in enumerate(company_reports_list):
                if item2 [0] == line2 [2]:
                    missing_reports [i2].insert(0, line2 [0])

        return (missing_reports, existing_reports)
    except Exception as e:
        restart(e, __name__)
def reportContent(company, r, row):
    try:
        global google
        if google != True:
            google = googleAPI()

        full_report = []
        report = ''

        print('...report for', row [1], row [2], row [3], row [4])
        # company
        # url = row[0]
        # relat = row[2]
        # optA = row[3]
        # optB = row[4]

        # RAD GRABBER
        if 'rad.cvm.gov.br' in row [0]:
            # RAD-DRE GRABBER
            if 'Dados' in row [2]:
                report = reportContentRAD_Dados(company, row)
            elif 'DRE' in row [2]:
                if DRE != row[3]:
                    report = reportContentRAD_DRE(company, row)
        if report:
            updateReportToSheet(company, report)
            full_report.extend(report)
            print('...saved')
        else:
            print('...not found')


    except Exception as e:
        restart(e, __name__)



# System Warp-Up/Cool Down Block
def start():
    try:
        global timestamp
        global google

        print('-- Hey Ho,')
        print('Let\'s Go!')
        timestamp = datetime.now()
        timestamp = timestamp.strftime("%d/%m/%Y %H:%M:%S")
        print(timestamp)

        # start engines
        engine = startEngine()

        # start Google
        google = googleAPI()
    except Exception as e:
        restart(e, __name__)
def end():
    try:
        print('-- This is the end, my friend')
        browser.quit()
    except Exception as e:
        print('-- Ops, Sheet Happens!')
        quit()
def restart(e, msg):
    try:
        browser.quit()
        print('stop working in', msg, e)
        allinone_project()
        quit()
    except:
        print('Erro terminal desconhecido. A coisa foi grave!')
        browser.quit()
        quit()


# Projects Block
def allinone_project():
    try:
        print('ALL-IN-ONE PROJECT in action')
        a = user_defined_variables()
        b = start()
        c = companyList()
        d = companyOrder()

        for i, company in enumerate(d):
            if i < batch_companies:
                company = companySheet(company)
                company_reports_list = companyReportList(company)
                reports = companyReportListData(company, company_reports_list)
                for r, row in enumerate(reports[0]):
                    if r < batch_reports:
                        missing_reports = reportContent(company, r, row)

        # z = end()
    except Exception as e:
        restart(e, __name__)


# run_in_a_line_geek
# aa = summer_project()  # list of companies
# bb = vacation_project()  # company sheet and data
# cc = winter_project()  # reports content
# dd = spring_project() # datastudio
ee = allinone_project()  # all in one

z = end()
