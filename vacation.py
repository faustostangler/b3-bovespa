# -*- encoding: utf-8 -*-
import varname
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
# updateCompanies
import time
# fase 2 imports
import re
# fase 2.1/2.2 imports
from datetime import datetime
from datetime import timedelta
# fase 3 imports
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# grab report imports
from selenium.webdriver.support.ui import Select

# USER DEFINED Variables
def user_defined_variables():
    try:
        # main page URL
        global main_url
        main_url = 'http://bvmf.bmfbovespa.com.br/cias-listadas/empresas-listadas/BuscaEmpresaListada.aspx?idioma=pt-br'

        # ID of Google Sheet to store list of companies
        global index_sheet
        index_sheet = '1Nc30KAFZ2jDGe7-uReGUGBT3yXDgzwUqweyfFeQlfGg'

        # ID of Google Sheet Template for new company sheets
        global template_sheet
        global default_sheets
        template_sheet = '1wsaqwT-WX_UhsKRrbgqivBc1S3KvKFGjq2MaW4e9ZxI'
        default_sheets = ['INDEX', 'F-WEB', 'F', 'Glossário', 'blasterlista']

        # google API factor of true
        global google
        google = False

        # google_sheet API authorization - see https://console.developers.google.com/
        global CLIENT_SECRET_FILE
        global ACCOUNT_SECRET_FILE
        global SHEET_SCOPE
        CLIENT_SECRET_FILE = 'client_secret.json'
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
        batch_reports = 20

    except Exception as e:
        restart(e, __name__)


# MAIN Block
def getListOfCompanies():
    try:
        global list_of_companies
        list_of_companies = []

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
                                            r) + ']/td[1]/a').get_attribute('href').split("=")[1],
                   browser.find_element_by_xpath(
                       '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[' + str(
                           r) + ']/td[2]').text, browser.find_element_by_xpath(
                    '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[' + str(
                        r) + ']/td[1]').text, browser.find_element_by_xpath(
                    '//*[@id="ctl00_contentPlaceHolderConteudo_BuscaNomeEmpresa1_grdEmpresa_ctl01"]/tbody/tr[' + str(
                        r) + ']/td[3]').text]
            list_of_companies.append(row)
            if r % 50 == 0:
                print(r)
        print(rowsA)
        print('...done')
    except Exception as e:
        restart(e, __name__)
def getSheetOfCompanies():
    try:
        if google != True:
            googleAPI()

        print('LOAD sheet of companies')
        # get planilha all values
        try:
            sheet_of_companies = []
            sheet_of_companies = ws_bovespa_lista_bovespa.get_all_values()

            sheet_of_companies.remove(sheet_of_companies[0])
        except:
            print('sheet_of_companies =', sheet_of_companies)
        for item in sheet_of_companies:
            item.pop()
            item.pop()

        not_found = list_difference(sheet_of_companies, list_of_companies)
        exceptions = [['14826', 'P.ACUCAR-CBD', 'CIA BRASILEIRA DE DISTRIBUICAO', 'NM'],['14826', 'P.ACUCAR-CBD', 'CIA BRASILEIRA DE DISTRIBUICAO', 'N1']]
        not_found = list_difference(not_found, exceptions)
        if not_found:
            # batch add list
            sheet = ws_bovespa_lista_bovespa
            lista = sheet.get_all_values()
            data = not_found
            start_col = 1
            start_row = len(lista) + 1
            # grid table math to crate sheet_range
            sheet_row = len(lista)
            sheet_col = len(lista[0]) # or 6??
            data_row = len(data)
            data_col = len(data[0])
            total_rows = max((start_row + data_row), sheet_row)
            total_cols = start_col + max(sheet_col, data_col) - 1

            sheet_range1 = sheetCol(start_col) + str(start_row)
            sheet_range2 = sheetCol(data_col) + str(total_rows)
            sheet_range = sheet_range1 + ':' + sheet_range2
            sheet.resize(total_rows, total_cols)
            print('resize1', sheet, total_rows, total_cols)

            cell_list = sheet.range(sheet_range)
            try:
                for cell in cell_list:
                    cell.value = data[cell.row - start_row][cell.col - 1]
                    # print(cell.row - start_row, cell.col -1)
            except:
                pass
            sheet.update_cells(cell_list)
            print('ws_bovespa_lista_bovespa 1 xxxfsxxx = ', ws_bovespa_lista_bovespa)
            quit()

        print(len(not_found), 'Companies updated to sheet')
        print('...done')
    except Exception as e:
        restart(e, __name__)
def sortListOfCompaniesByCompany():
    try:
        global list_of_companies

        if google != True:
            googleAPI()

        print('ORDER BY list of companies')
        list_of_companies = ws_bovespa_lista_bovespa.get_all_values()
        list_of_companies.remove(list_of_companies[0])
        for item in list_of_companies:
            item.pop()

        # insert provisory timestamp
        timestamp2 = datetime.now() - timedelta(days=7)
        timestamp2 = timestamp2.strftime("%d/%m/%Y %H:%M:%S")

        for i in list_of_companies:
            try:
                a = i[4]
                if not i[4]:
                    i[4] = timestamp2
            except:
                i.append(timestamp2)

        # order by timestamp newer
        try:
            list_of_companies = sorted(list_of_companies, key=lambda x: (x[1]))
            list_of_companies = sorted(list_of_companies,
                                       key=lambda x: datetime.strptime(x[4], "%d/%m/%Y %H:%M:%S"))
        except:
            pass
        # remove timestamp col
        for i in list_of_companies:
            try:
                del i[4]
            except:
                pass
        return list_of_companies
        print('...done')
    except Exception as e:
        restart(e, __name__)
def sortListOfCompaniesByReport():
    try:
        global list_of_companies

        if google != True:
            googleAPI()

        print('ORDER BY list of companies reports')
        list_of_companies = ws_bovespa_lista_bovespa.get_all_values()
        list_of_companies.remove(list_of_companies[0])
        # for item in list_of_companies:
        #     item.pop()



        # insert provisory timestamp
        timestamp2 = datetime.now() - timedelta(days=7)
        timestamp2 = timestamp2.strftime("%d/%m/%Y %H:%M:%S")

        for i in list_of_companies:
            try:
                a = i[5]
                if not i[5]:
                    i[5] = timestamp2
            except:
                i.append(timestamp2)

        # order by timestamp newer
        try:
            list_of_companies = sorted(list_of_companies, key=lambda x: (x[1]))
            list_of_companies = sorted(list_of_companies,
                                       key=lambda x: datetime.strptime(x[5], "%d/%m/%Y %H:%M:%S"))
        except:
            pass
        # remove timestamp col
        for i in list_of_companies:
            try:
                i.pop()
            except:
                pass
        return list_of_companies
        print('...done')
    except Exception as e:
        restart(e, __name__)
def getCompany(company):
    try:
        print('GET INFO for', company[1])

        getCompanyMainPage(company)
        getCompanyListOfReports(company)

    except Exception as e:
        restart(e, __name__)
def updateCompanyToSheet(company):
    try:
        if google != True:
            googleAPI()

        # create or update sheet
        global newsheet_id

        print('... Google Sheets')
        # Get Company Sheet
        newsheet_id = googleGetSheet(index_sheet)

        vacationLog(company, 5)

        # google worksheets for company sheet
        sh_company = gsheet.open_by_key(newsheet_id)
        sh_company_list = sh_company.worksheets()
        ws_company_index = sh_company.worksheet('INDEX')
        ws_company_glossario = sh_company.worksheet('Glossário')
        ws_company_f = sh_company.worksheet('F')
        ws_company_fweb = sh_company.worksheet('F-WEB')
        try: # one can safely delete all this blasterlista try
            ws_company_blasterlista = sh_company.worksheet('blasterlista')
        except:
            pass
        # update company sheet
        ws_company_index.resize(6, 6)
        ws_company_index.update_acell('B3', company[0])
        ws_company_index.update_acell('B1', company[1])
        ws_company_index.update_acell('B2', company[2])
        ws_company_index.update_acell('E5', company[3])
        ws_company_index.update_acell('F5', company[4])
        ws_company_index.update_acell('B5', " ".join(company[5]))
        ws_company_index.update_acell('B4', company[6])
        ws_company_index.update_acell('C1', company[7])
        ws_company_index.update_acell('D2', company[8])
        ws_company_index.update_acell('D3', company[9])
        ws_company_index.update_acell('D4', company[10])
        ws_company_index.update_acell('D5', company[11])
        ws_company_index.update_acell('E3', company[12])
        ws_company_index.update_acell('E4', company[13])
        print('... done cover')

        # batch add list of reports
        sheet = ws_company_index
        lista = sheet.get_all_values()
        data = company[14]
        start_col = 1
        start_row = 7
        # grid table math to crate sheet_range
        sheet_row = len(lista)
        sheet_col = len(lista[0])
        data_row = len(data)
        data_col = len(data[0])
        total_rows = max((start_row + data_row), sheet_row)
        total_cols = start_col + max(sheet_col, data_col) - 1

        sheet_range1 = sheetCol(start_col) + str(start_row)
        sheet_range2 = sheetCol(data_col) + str(total_rows)
        sheet_range = sheet_range1 + ':' + sheet_range2
        sheet.resize(total_rows, total_cols)
        cell_list = sheet.range(sheet_range)
        try:
            for cell in cell_list:
                cell.value = data[cell.row - start_row][cell.col - 1]
        except:
            pass
        sheet.update_cells(cell_list)
        print('... done list of reports')

        print('...done')

    except Exception as e:
        print('Please wait 60 seconds for Google API to recover from quota exhaustion')
        time. sleep(60)
        restart(e, __name__)
def getSheetListOfReports(c):
    try:
        global full_report

        global newsheet_id
        global company
        company = c



        if google != True:
            googleAPI()

        print('GET REPORTS for', company[1])
        # Get Existing Company Sheet newsheet_id
        newsheet_id = ''
        sheet = ws_bovespa_uberlista
        list = sheet.get_all_values()
        cvm_url = 'http://bvmf.bmfbovespa.com.br/pt-br/mercados/acoes/empresas/ExecutaAcaoConsultaInfoEmp.asp?CodCVM='
        sheet_url = 'https://docs.google.com/spreadsheets/d/'
        if newsheet_id == '':
            try:
                search = str(company [0])  # item to search for (CVM code)
                for r, row in enumerate(list):
                    if search == row [2].replace(cvm_url, ''):  # second index/third column in list
                        newsheet_id = sheet.cell(r + 1, 4).value
                        newsheet_id = newsheet_id.replace(sheet_url, '')
                        # update company from uberlista, just for fun
                        company[4] = sheet.cell(r + 1, 5).value
                        company.append(sheet.cell(r + 1, 9).value.split('xxxfsxxx'))
                        company.append(sheet.cell(r + 1, 8).value)
                        company.append(sheet.cell(r + 1, 16).value)
                        company.append(sheet.cell(r + 1, 13).value)
                        company.append(sheet.cell(r + 1, 14).value)
                        company.append(sheet.cell(r + 1, 15).value)
                        company.append(sheet.cell(r + 1, 6).value)
                        company.append(sheet.cell(r + 1, 10).value)
                        company.append(sheet.cell(r + 1, 11).value)
                        action = 'found'
                        break
            except:
                    pass

        if newsheet_id == '':
            name = {'name': str(company[5][0][:4].replace('Nenh','NONE') + ' - ' + company[1] + ' - ' + company[2] + ' - ' + company[0])}
            newsheet = gdrive.files().copy(fileId=template_sheet, body=name).execute()
            newsheet_id = newsheet.get('id')
            permission_user = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': CLIENT_EMAIL
            }
            print('PLEASE REMEMBER TO REACTIVATE PERMISSIONS IN 24h')
            # permission = gdrive.permissions().create(fileId=newsheet_id, body=permission_user, fields="id").execute()
            # permission_id = permission.get('id')
            action = 'created'
        print(action, company[5][0][:4].replace('Nenh','NONE'), company[0], newsheet_id)

        # google worksheets for company sheet
        sh_company = gsheet.open_by_key(newsheet_id)
        sh_company_list = sh_company.worksheets()
        ws_company_index = sh_company.worksheet('INDEX')
        ws_company_glossario = sh_company.worksheet('Glossário')
        ws_company_f = sh_company.worksheet('F')
        ws_company_fweb = sh_company.worksheet('F-WEB')
        try: # one can safely delete all this blasterlista try
            ws_company_blasterlista = sh_company.worksheet('blasterlista')
        except:
            pass

        list_of_reports = ws_company_index.get_all_values()
        list_of_reports = [line for line in list_of_reports if 'http' in line [0]]
        total_reports = [line[2] for line in list_of_reports if 'http' in line [0]]

        sheet_reports = [sheet.title for sheet in sh_company_list]

        existing_reports = list_difference(sheet_reports, default_sheets)
        missing_reports = list_difference(total_reports, existing_reports)
        if not missing_reports:
            missing_reports = total_reports

        # reset to start grabbing reports from beggining
        # if existing_reports:
        #     for i, r in enumerate(existing_reports):
        #         missing_reports.insert(i, r)


        list_of_reports = [row for row in list_of_reports if row [2] in missing_reports]


        for r, row in enumerate(list_of_reports):
            if r < batch_reports:
                full_report = []
                if 'http' in row [0]:
                    # RAD STYLE
                    if 'rad.cvm.gov.br' in row [0]:
                        print('loading report', row [2])
                        # # get report11
                        # relat = 'DRE'
                        # optA = 'Dados da Empresa'
                        # optB = 'Composição do Capital'
                        # txt = 'stock'
                        # href = row[0]
                        # print('DO IT IN LAST PLACE')
                        # # ReportGeneratorRAD(optA, optB, relat, txt, href)
                        # print(relat, optA, optB, '******************************* DO IT IN LAST PLACE 11 ************************')
                        # print('report', report)

                        optA = "DFs Individuais"
                        reportGrupo(row, optA)

                        optA = "DFs Consolidadas"
                        reportGrupo(row, optA)

                        report_sheet_id = reportToSheet(row)
                        createBlasterlista(company)
                        winterLog(company, row [2], report_sheet_id)
                        print(r+1, 'Reports updated to sheet',  row [2])
                        print('https://docs.google.com/spreadsheets/d/'+newsheet_id)

                    # BMFBOVESPA STYLE
                    if 'bmfbovespa.com.br' in row [0]:
                        print('skipped report for', row [2], row [0])

                    print(row [0])


    except Exception as e:
        restart(e, __name__)
def reportGrupo(row, optA):
    try:
        global xpathA

        browser.get(row[0])
        xpathA = '//*[@id="cmbGrupo"]'
        browser.find_element(By.XPATH, xpathA)

        wait.until(EC.presence_of_element_located((By.XPATH, xpathA)))
        selectA = Select(browser.find_element(By.XPATH, xpathA))
        selectA.select_by_visible_text(optA)
        # time.sleep(1)
        # browser.minimize_window()

        reportQuadro(row, optA)
    except:
        print(optA, '...not found')
def reportQuadro(row, optA):
    try:
        relat = "DRE"
        href = row [0]

        ### optB REPORT
        # report1
        optB = "Balanço Patrimonial Ativo"
        txt = "cell"
        ReportGeneratorRAD(optA, optB, relat, txt, href, row [2])

        # report2
        optB = "Balanço Patrimonial Passivo"
        txt = "cell"
        ReportGeneratorRAD(optA, optB, relat, txt, href, row [2])

        # report3
        optB = "Demonstração do Resultado"
        txt = "txt"
        ReportGeneratorRAD(optA, optB, relat, txt, href, row [2])

        # report4
        optB = "Demonstração do Resultado Abrangente"
        href = row [0]
        ReportGeneratorRAD(optA, optB, relat, txt, href, row [2])

        # report5
        optB = "Demonstração do Fluxo de Caixa"
        txt = "txt"
        ReportGeneratorRAD(optA, optB, relat, txt, href, row [2])

        # report6
        optB = "Demonstração de Valor Adicionado"
        txt = "txt"
        ReportGeneratorRAD(optA, optB, relat, txt, href, row [2])
    except Exception as e:
        restart(e, __name__)
def updateUberblasterlista():
    try:
        if google != True:
            googleAPI()

        ws_bovespa_uberblasterlista = sh_bovespa.worksheet('uberblasterlista')
        ws_bovespa_uberblasterlista.resize(1, 8)
        ws_bovespa_uberblasterlista.resize(2, 8)
        sheet_url = 'https://docs.google.com/spreadsheets/d/'

        newsheet_id = [item.replace(sheet_url, '') for item in ws_bovespa_uberlista.col_values(4) if sheet_url in item]

        if newsheet_id:
            prepre = '=SORT(UNIQUE({'
            pospos = '});8;TRUE;7;TRUE;4;TRUE;5;TRUE;1;TRUE)'
            pre = 'IFERROR(IMPORTRANGE("'
            pos = '";"blasterlista!A2:H"))'
            joint = '; '

            formula = pos + joint + pre
            formula = formula.join(newsheet_id)
            formula = prepre + pre + formula + pos + pospos
            ws_bovespa_uberblasterlista.update_acell('A2', formula)


        print('... done')




    except Exception as e:
        restart(e, __name__)
def vacationLog(c, col):
    try:
        if google != True:
            googleAPI()
        company = c


        cvm_url = 'http://bvmf.bmfbovespa.com.br/pt-br/mercados/acoes/empresas/ExecutaAcaoConsultaInfoEmp.asp?CodCVM='
        sheet_url = 'https://docs.google.com/spreadsheets/d/'

        print('... log company ', company[1])
        # update ws_bovespa_log
        right_now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        log_company = [str(company [0]), str(company [5] [0] [:4].replace('Nenh', 'NONE')), sheet_url + newsheet_id, right_now, 'company']
        ws_bovespa_log.append_row(log_company)

        # update ws_bovespa_lista_bovespa
        cell = ws_bovespa_lista_bovespa.find(company[0])
        row = cell.row
        ws_bovespa_lista_bovespa.update_cell(row, col, datetime.now().strftime("%d/%m/%Y %H:%M:%S"))


        # update ws_bovespa_uberlista
        log_uberlista = [company[1], company[0], cvm_url + company[0], sheet_url + newsheet_id, company[4], company[11], company[2], company[6], " ".join(company[5]), company[12], company[13], company[3], company[8], company[9], company[10], company[7]]
        sheet = sh_bovespa.worksheet('uberlista')
        col = sheet.col_count
        row = sheet.row_count+1
        try:
            cell = sheet.find(company [0])
            row = cell.row
        except:
            sheet.resize(row, col)
        start_row = row
        start_col = 1

        sheet_range1 = sheetCol(start_col) + str(start_row)
        sheet_range2 = sheetCol(col) + str(start_row)
        sheet_range = sheet_range1 + ':' + sheet_range2

        cell_list = sheet.range(sheet_range)
        try:
            for cell in cell_list:
                cell.value = log_uberlista[cell.col - 1]
        except:
            pass
        sheet.update_cells(cell_list)




        print('... done')
    except Exception as e:
        restart(e, __name__)
def winterLog(c, report, report_sheet_id):
    try:
        if google != True:
            googleAPI()
        company = c
        sheet_id = report_sheet_id

        print('... log company ', company[1])
        # update log sheet
        right_now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        log_url = 'https://docs.google.com/spreadsheets/d/' + newsheet_id + '/edit#gid=' + str(sheet_id)
        log_data = [str(company[0]), str(company[5][0][:4].replace('Nenh','NONE')), log_url, right_now, 'report ' + report]
        ws_bovespa_log.append_row(log_data)
        print('...log done!', company[1])
    except Exception as e:
        restart(e, __name__)

# Background Block
def googleAPI():
    try:
        global google
        google = True

        print('GOOGLE API authorization')
        global gdrive
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

        # sheet worksheets
        global sh_bovespa
        global ws_bovespa_uberblasterlista
        global ws_bovespa_log
        global ws_bovespa_lista_bovespa
        global ws_bovespa_uberlista
        sh_bovespa = gsheet.open_by_key(index_sheet)
        ws_bovespa_uberblasterlista = sh_bovespa.worksheet('uberblasterlista')
        ws_bovespa_log = sh_bovespa.worksheet('log')
        ws_bovespa_lista_bovespa = sh_bovespa.worksheet('lista_bovespa')
        ws_bovespa_uberlista = sh_bovespa.worksheet('uberlista')
        print('...done')

    except Exception as e:
        print('google API overload... wait 1 minute')
        time.sleep(60)
        googleAPI()
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
        range1 = str(sheetCol(start_col))+str(start_row)
        range2 = str(sheetCol(start_col+len(data[0])-1))+str(len(data)+start_row-1)
        sheet_range = range1 + ':' + range2
        return sheet_range
    except Exception as e:
        restart(e, __name__)
def getCompanyMainPage(c):
    try:
        global company
        company = c
        print('... Company Details')
        browser.get(
            'http://bvmf.bmfbovespa.com.br/pt-br/mercados/acoes/empresas/ExecutaAcaoConsultaInfoEmp.asp?CodCVM=' + company[
                0])
        browser.minimize_window()

        wait.until(EC.presence_of_element_located((By.XPATH, '/html/body')))
        # update
        update = getCompanyItem('/html/body/', 'div[1]/div[1]')
        update = update.replace('Atualizado em ', '').replace(', às', '')
        update = datetime.strptime(update, '%d/%m/%Y %Hh%M')
        update = update.strftime("%d/%m/%Y %H:%M:00")
        company.append(update)
        # ticker
        ticker = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[2]/td[2]')
        tickers = re.split(' ', ticker.replace('Mais Códigos\n','').replace(';',''))
        company.append(tickers)
        # cnpj
        cnpj = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[3]/td[2]')
        company.append(cnpj)
        # atividade
        atividade = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[4]/td[2]')
        company.append(atividade)
        # setor
        setor = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[5]/td[2]')
        setores = re.split(' / ', setor)
        company.append(setores[0])
        company.append(setores[1])
        company.append(setores[2])
        # site
        site = getCompanyItem('/html/body/div[2]/div[1]/ul/li[1]/div', '/table/tbody/tr[6]/td[2]')
        company.append(site)
        # skipped for now
        # # on
        # on = getCompanyItem('/html/body/div[3]/div/div/div/div/div[2]/div[3]/div', '/table/tbody/tr[1]/td[2]')
        # company.append(on)
        # # pn
        # pn = getCompanyItem('//html/body/div[3]/div/div/div/div/div[2]/div[3]/div', '/table/tbody/tr[2]/td[2]')
        # company.append(pn)
        company.append('on skipped')
        company.append('pn skipped')
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
def getCompanyListOfReports(c):
    try:
        global company
        company = c

        company_reports = []

        dfp = getReportList(company, 'dfp')
        company_reports.extend(dfp)
        itr = getReportList(company, 'itr')
        company_reports.extend(itr)

        company_reports = sorted(company_reports, key=lambda x: datetime.strptime(x[2], "%d/%m/%Y"), reverse=True)
        for i, item in enumerate(company_reports):
            data = datetime.strptime(item[2], '%d/%m/%Y')
            company_reports[i][2] = data.strftime('%Y') + data.strftime('%m') + data.strftime('%d')
        company.append(company_reports)
    except Exception as e:
        restart(e, __name__)
def getReportList(c, report):
    try:
        global company
        company = c

        print('... Company Reports', report)
        link = 'http://bvmf.bmfbovespa.com.br/cias-listadas/empresas-listadas/HistoricoFormularioReferencia.aspx?codigoCVM=' + \
              company[0] + '&tipo=' + report + '&ano=0&idioma=pt-br'
        # load company report page
        browser.get(link)
        browser.minimize_window()

        assert (EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_contentPlaceHolderConteudo_divDemonstrativo"]')))
        wait.until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_contentPlaceHolderConteudo_divDemonstrativo"]')))
        links = browser.find_elements(By.TAG_NAME, 'a')
        list_of_reports = reportLinkParser(links)
        return list_of_reports
    except Exception as e:
        restart(e, __name__)
def reportLinkParser(links):
    try:
        list_of_reports = []
        pub_date = ''
        for link in links:
            # check if report is newer
            text = str(link.text)
            href = str(link.get_attribute('href'))
            if text:
                if pub_date != text.split(" - ")[0]:
                    # parse new style links
                    if 'AbreFormularioCadastral' in href:
                        report = href.replace("&CodigoTipoInstituicao=2')", "").replace(
                            "javascript:AbreFormularioCadastral('http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=",
                            "")
                        link = href.replace("')", "").replace("javascript:AbreFormularioCadastral('", "")
                        report_item = [link, report, text.split(" - ")[0], text.split(" - ")[1],
                                       text.split(" - ")[2]]
                        list_of_reports.append(report_item)
                    # parse old style links
                    if 'ConsultarDXW' in href:
                        report = href.replace("')", "").replace(
                            "javascript:ConsultarDXW('http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?", "")
                        link = href.replace("')", "").replace("javascript:ConsultarDXW('", "")
                        report_item = [link, company[0], text.split(" - ")[0], text.split(" - ")[1],
                                       text.split(" - ")[2]]
                        list_of_reports.append(report_item)
                pub_date = text.split(" - ")[0]
        return list_of_reports
    except Exception as e:
        restart(e, __name__)
def googleGetSheet(index_sheet):
    global newsheet_id
    try:
        newsheet_id = ''
        action = 'unkown'
        # if sheet in ws_bovespa_uberlista, get code, or create
        # Get Existing Company Sheet newsheet_id
        sheet = ws_bovespa_uberlista
        list = sheet.get_all_values()
        cvm_url = 'http://bvmf.bmfbovespa.com.br/pt-br/mercados/acoes/empresas/ExecutaAcaoConsultaInfoEmp.asp?CodCVM='
        sheet_url = 'https://docs.google.com/spreadsheets/d/'
        if newsheet_id == '':
            search = str(company[0])  # item to search for (CVM code)
            for r, row in enumerate(list):
                if search == row[2].replace(cvm_url, ''):  # second index/third column in list
                    newsheet_id = sheet.cell(r + 1, 4).value
                    newsheet_id = newsheet_id.replace(sheet_url, '')
                    action = 'found'
                    break
        if newsheet_id == '':
            name = {'name': str(company[5][0][:4].replace('Nenh','NONE') + ' - ' + company[1] + ' - ' + company[2] + ' - ' + company[0])}
            newsheet = gdrive.files().copy(fileId=template_sheet, body=name).execute()
            newsheet_id = newsheet.get('id')
            permission_user = {
                'type': 'user',
                'role': 'writer',
                'emailAddress': CLIENT_EMAIL
            }
            print('PLEASE REMEMBER TO REACTIVATE PERMISSIONS IN 24h')
            # permission = gdrive.permissions().create(fileId=newsheet_id, body=permission_user, fields="id").execute()
            # permission_id = permission.get('id')
            action = 'created'
        print(action, company[5][0][:4].replace('Nenh','NONE'), company[0], newsheet_id)
        return newsheet_id
    except Exception as e:
        restart(e, __name__)
def ReportGeneratorRAD(optA, optB, relat, txt, href, report):
    try:
        # browser.get(href)
        # browser.minimize_window()
        # time.sleep(1)
        xpathB = '//*[@id="cmbQuadro"]'
        frame_name = 'iFrameFormulariosFilho'
        frame = '//*[@id="iFrameFormulariosFilho"]'
        table = '//*[@id="ctl00_cphPopUp_tbDados"]'
        browser.switch_to.default_content()

        # browser.switch_to.window(browser.current_window_handle)
        # browser.find_element(By.XPATH, xpathA)
        browser.find_element(By.XPATH, xpathB)

        wait.until(EC.presence_of_element_located((By.XPATH, xpathB)))
        selectB = Select(browser.find_element(By.XPATH, xpathB))
        selectB.select_by_visible_text(optB)
        # browser.minimize_window()


        time.sleep(1)
        wait.until(EC.presence_of_element_located((By.XPATH, frame)))
        browser.switch_to.frame(frame_name)
        # browser.minimize_window()

        # specific conditions
        if txt == 'stock':
            table = '//*[@id="UltimaTabela"]/table'
            on = '//td[@id="QtdAordCapiItgz_1"]'
            pn = '//td[@id="QtdAprfCapiItgz_1"]'

        wait.until(EC.presence_of_element_located((By.XPATH, table)))
        rows: int = len(browser.find_elements(By.XPATH, '//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr'))
        cols: int = len(browser.find_elements(By.XPATH, '//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[1]/td'))

        row = []

        for r in range(1, rows +1):
            col = []
            col1 = browser.find_element_by_xpath('//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[' + str(r) + ']/td[1]').text
            col2 = browser.find_element_by_xpath('//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[' + str(r) + ']/td[2]').text
            col3 = browser.find_element_by_xpath('//*[@id="ctl00_cphPopUp_tbDados"]/tbody/tr[' + str(r) + ']/td[3]').text
            if '\n a \n' in col3:
                col3 = col3.split('\n')[2] + ' -- ' + col3
            col.append(col1)
            col.append(col2)
            col.append(col3)
            col.append(optA)
            col.append(optB)
            col.append(relat)
            col.append(report)
            col.append(company[1])
            row.append(col)
        full_report.append(row)
        print(optA, ' - ', optB, '...done')
    except:
        print(optA, ' - ', optB, '...not found')
def reportToSheet(row):
    try:
        if google != True:
            googleAPI()

        # delete present worksheet row[2]
        sh_company = gsheet.open_by_key(newsheet_id)
        sh_company_list = sh_company.worksheets()
        ws_company_index = sh_company.worksheet('INDEX')
        ws_company_glossario = sh_company.worksheet('Glossário')
        ws_company_f = sh_company.worksheet('F')
        ws_company_fweb = sh_company.worksheet('F-WEB')
        try: # one can safely delete all this blasterlista try
            ws_company_blasterlista = sh_company.worksheet('blasterlista')
        except:
            pass

        worksheets = [sheet for sheet in sh_company_list]
        worksheets_title = [sheet.title for sheet in sh_company_list]
        try:
            n = worksheets_title.index(row[2])
            sh_company.del_worksheet(worksheets[n])
            print('... cleared', row [2])
        except:
            pass

        # prepare full_report
        sh_company = gsheet.open_by_key(newsheet_id)
        # remove title line from multiple reports inside full_report
        for n in range(0, len(full_report)):
            # print(n, full_report[n][0])
            if len(full_report[n]) > 0:
                del full_report[n][0]
            else:
                pass
        # prepare flat report
        flat_report = []
        for i, report in enumerate(full_report):
            for j, line in enumerate(report):
                flat_report.append(line)
        # insert title at top
        flat_report.insert(0,[' Conta ', ' Descrição ', 'Resultado', 'Grupo', 'Quadro', 'DRE', 'Empresa', 'Relatório'])

        # worksheet creation
        ws_company_report = sh_company.add_worksheet(title=row[2], rows=len(flat_report), cols=len(flat_report[0]))
        global report_sheet_id
        report_sheet_id = ws_company_report.id

        # batch cell update full_report
        data = flat_report
        start_col = 1
        start_row = 1
        sheet = ws_company_report
        range1 = str(sheetCol(start_col))+str(start_row)
        range2 = str(sheetCol(start_col+len(data[0])-1))+str(len(data)+start_row-1)
        sheet_range = range1 + ':' + range2
        cell_list = sheet.range(sheet_range)
        for cell in cell_list:
            cell.value = data[cell.row-1][cell.col-1]
        sheet.update_cells(cell_list)
        return report_sheet_id
    except Exception as e:
        restart(e, __name__)
def createBlasterlista(c):
    try:
        global company
        company = c

        if google != True:
            googleAPI()

        # create or update sheet
        global newsheet_id

        # google worksheets for company sheet
        sh_company = gsheet.open_by_key(newsheet_id)
        sh_company_list = sh_company.worksheets()
        # ws_company_index = sh_company.worksheet('INDEX')
        # ws_company_glossario = sh_company.worksheet('Glossário')
        # ws_company_f = sh_company.worksheet('F')
        # ws_company_fweb = sh_company.worksheet('F-WEB')
        try:
            ws_company_blasterlista = sh_company.worksheet('blasterlista')
        except:
            pass

        # prepare list of sheets to formula
        sheets = sh_company.worksheets()

        sheet_reports = [sheet.title for sheet in sheets]
        existing_reports = list_difference(sheet_reports, default_sheets)
        existing_reports.sort(reverse=True)

        if existing_reports:
            pos = '!A2:H")'
            pre = 'INDIRECT("'
            formula = pos + '; ' + pre
            formula = formula.join(existing_reports)
            formula = '={' + pre + formula + pos + '}'
            ws_company_blasterlista.update_acell('A2', formula)
    except Exception as e:
        restart(e, __name__)


# DEBUG Block
def getListOfCompaniesDEBUG():
    try:
        global list_of_companies
        print('LOAD page DEBUG')
        list_of_companies = [['16284', '524 PARTICIP', '524 PARTICIPACOES S.A.', 'MB'],['21725', 'ADVANCED-DH', 'ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.', ' '],['18970', 'AES TIETE E', 'AES TIETE ENERGIA SA', 'N2'],['22179', 'AFLUENTE T', 'AFLUENTE TRANSMISSÃO DE ENERGIA ELÉTRICA S/A', ' '],['16705', 'ALEF S/A', 'ALEF S.A.', 'MB'],['9954', 'ALFA HOLDING', 'ALFA HOLDINGS S.A.', ' '],['21032', 'ALGAR TELEC', 'ALGAR TELECOM S/A', ' '],['22357', 'ALIANSCSONAE', 'ALIANSCE SONAE SHOPPING CENTERS S.A.', 'NM'],['10456', 'ALPARGATAS', 'ALPARGATAS S.A.', 'N1'],['22217', 'ALPER S.A.', 'ALPER CONSULTORIA E CORRETORA DE SEGUROS S.A.', 'NM'],['18066', 'ALTERE SEC', 'ALTERE SECURITIZADORA S.A.', ' '],['21490', 'ALUPAR', 'ALUPAR INVESTIMENTO S/A', 'N2'],['23264', 'AMBEV S/A', 'AMBEV S.A.', ' '],['3050', 'AMPLA ENERG', 'AMPLA ENERGIA E SERVICOS S.A.', ' '],['23248', 'ANIMA', 'ANIMA HOLDING S.A.', 'NM'],['22349', 'AREZZO CO', 'AREZZO INDÚSTRIA E COMÉRCIO S.A.', 'NM'],['24171', 'CARREFOUR BR', 'ATACADÃO S.A.', 'NM'],['15423', 'ATOMPAR', 'ATOM EMPREENDIMENTOS E PARTICIPAÇÕES S.A.', ' '],['11975', 'AZEVEDO', 'AZEVEDO E TRAVASSOS S.A.', ' '],['24112', 'AZUL', 'AZUL S.A.', 'N2'],['20990', 'B2W DIGITAL', 'B2W - COMPANHIA DIGITAL', 'NM'],['21610', 'B3', 'B3 S.A. - BRASIL. BOLSA. BALCÃO', 'NM'],['701', 'BAHEMA', 'BAHEMA EDUCAÇÃO S.A.', 'MA'],['24600', 'BANCO BMG', 'BANCO BMG S.A.', 'N1'],['24406', 'BANCO INTER', 'BANCO INTER S.A.', 'N2'],['1155', 'BANESTES', 'BANESTES S.A. - BCO EST ESPIRITO SANTO', ' '],['1520', 'BARDELLA', 'BARDELLA S.A. INDUSTRIAS MECANICAS', ' '],['15458', 'BATTISTELLA', 'BATTISTELLA ADM PARTICIPACOES S.A.', ' '],['1562', 'BAUMER', 'BAUMER S.A.', ' '],['23159', 'BBSEGURIDADE', 'BB SEGURIDADE PARTICIPAÇÕES S.A.', 'NM'],['24660', 'BBMLOGISTICA', 'BBM LOGISTICA S.A.', 'MA'],['20958', 'ABC BRASIL', 'BCO ABC BRASIL S.A.', 'N2'],['1384', 'ALFA INVEST', 'BCO ALFA DE INVESTIMENTO S.A.', ' '],['922', 'AMAZONIA', 'BCO AMAZONIA S.A.', ' '],['906', 'BRADESCO', 'BCO BRADESCO S.A.', 'N1'],['1023', 'BRASIL', 'BCO BRASIL S.A.', 'NM'],['22616', 'BTGP BANCO', 'BCO BTG PACTUAL S.A.', 'N2'],['1120', 'BANESE', 'BCO ESTADO DE SERGIPE S.A. - BANESE', ' '],['1171', 'BANPARA', 'BCO ESTADO DO PARA S.A.', ' '],['1210', 'BANRISUL', 'BCO ESTADO DO RIO GRANDE DO SUL S.A.', 'N1'],['20885', 'INDUSVAL', 'BCO INDUSVAL S.A.', 'N2'],['1309', 'MERC INVEST', 'BCO MERCANTIL DE INVESTIMENTOS S.A.', ' '],['1325', 'MERC BRASIL', 'BCO MERCANTIL DO BRASIL S.A.', ' '],['1228', 'NORD BRASIL', 'BCO NORDESTE DO BRASIL S.A.', ' '],['21199', 'BANCO PAN', 'BCO PAN S.A.', 'N1'],['20567', 'PINE', 'BCO PINE S.A.', 'N2'],['20532', 'SANTANDER BR', 'BCO SANTANDER (BRASIL) S.A.', ' '],['19747', 'BETA SECURIT', 'BETA SECURITIZADORA S.A.', ' '],['17884', 'BETAPART', 'BETAPART PARTICIPACOES S.A.', 'MB'],['1694', 'BIC MONARK', 'BICICLETAS MONARK S.A.', ' '],['19305', 'BIOMM', 'BIOMM S.A.', 'MA'],['22845', 'BIOSEV', 'BIOSEV S.A.', 'NM'],['80179', 'BIOTOSCANA', 'BIOTOSCANA INVESTMENTS S.A.', 'DR3'],['24317', 'BK BRASIL', 'BK BRASIL OPERAÇÃO E ASSESSORIA A RESTAURANTES SA', 'NM'],['16772', 'BNDESPAR', 'BNDES PARTICIPACOES S.A. - BNDESPAR', 'MB'],['12190', 'BOMBRIL', 'BOMBRIL S.A.', ' '],['19909', 'BR MALLS PAR', 'BR MALLS PARTICIPACOES S.A.', 'NM'],['19925', 'BR PROPERT', 'BR PROPERTIES S.A.', 'NM'],['19640', 'BRADESCO LSG', 'BRADESCO LEASING S.A. ARREND MERCANTIL', ' '],['18724', 'BRADESPAR', 'BRADESPAR S.A.', 'N1'],['21180', 'BR BROKERS', 'BRASIL BROKERS PARTICIPACOES S.A.', 'NM'],['20036', 'BRASILAGRO', 'BRASILAGRO - CIA BRAS DE PROP AGRICOLAS', 'NM'],['4820', 'BRASKEM', 'BRASKEM S.A.', 'N1'],['19720', 'BRAZIL REALT', 'BRAZIL REALTY CIA SECURIT. CRÉD. IMOBILIÁRIOS', ' '],['17922', 'BRAZILIAN FR', 'BRAZILIAN FINANCE E REAL ESTATE S.A.', ' '],['18759', 'BRAZILIAN SC', 'BRAZILIAN SECURITIES CIA SECURITIZACAO', ' '],['14206', 'BRB BANCO', 'BRB BCO DE BRASILIA S.A.', ' '],['20672', 'BRC SECURIT', 'BRC SECURITIZADORA S.A.', ' '],['16292', 'BRF SA', 'BRF S.A.', 'NM'],['19984', 'BRPR 56 SEC', 'BRPR 56 SECURITIZADORA CRED IMOB S.A.', ' '],['23817', 'BRQ', 'BRQ SOLUCOES EM INFORMATICA S.A.', 'MA'],['20133', 'BV LEASING', 'BV LEASING - ARRENDAMENTO MERCANTIL S.A.', ' '],['19119', 'CABINDA PART', 'CABINDA PARTICIPACOES S.A.', 'MB'],['22683', 'CACHOEIRA', 'CACHOEIRA PAULISTA TRANSMISSORA ENERGIA S.A.', 'MB'],['19135', 'CACONDE PART', 'CACONDE PARTICIPACOES S.A.', 'MB'],['2100', 'CAMBUCI', 'CAMBUCI S.A.', ' '],['24228', 'CAMIL', 'CAMIL ALIMENTOS S.A.', 'NM'],['17493', 'CAPITALPART', 'CAPITALPART PARTICIPACOES S.A.', 'MB'],['18821', 'CCR SA', 'CCR S.A.', 'NM'],['24848', 'CEA MODAS', 'CEA MODAS S.A.', 'NM'],['13854', 'CEMEPE', 'CEMEPE INVESTIMENTOS S.A.', ' '],['20303', 'CEMIG DIST', 'CEMIG DISTRIBUICAO S.A.', ' '],['20320', 'CEMIG GT', 'CEMIG GERACAO E TRANSMISSAO S.A.', ' '],['2437', 'ELETROBRAS', 'CENTRAIS ELET BRAS S.A. - ELETROBRAS', 'N1'],['2461', 'CELESC', 'CENTRAIS ELET DE SANTA CATARINA S.A.', 'N2'],['24058', 'ALLIAR', 'CENTRO DE IMAGEM DIAGNOSTICOS S.A.', 'NM'],['2577', 'CESP', 'CESP - CIA ENERGETICA DE SAO PAULO', 'N1'],['14826', 'P.ACUCAR-CBD', 'CIA BRASILEIRA DE DISTRIBUICAO', 'N1'],['16861', 'CASAN', 'CIA CATARINENSE DE AGUAS E SANEAM.-CASAN', ' '],['21393', 'CELGPAR', 'CIA CELG DE PARTICIPACOES - CELGPAR', ' '],['16616', 'CEG', 'CIA DISTRIB DE GAS DO RIO DE JANEIRO-CEG', ' '],['14524', 'COELBA', 'CIA ELETRICIDADE EST. DA BAHIA - COELBA', ' '],['14451', 'CEB', 'CIA ENERGETICA DE BRASILIA', ' '],['2453', 'CEMIG', 'CIA ENERGETICA DE MINAS GERAIS - CEMIG', 'N1'],['14362', 'CELPE', 'CIA ENERGETICA DE PERNAMBUCO - CELPE', ' '],['14869', 'COELCE', 'CIA ENERGETICA DO CEARA - COELCE', ' '],['18139', 'COSERN', 'CIA ENERGETICA DO RIO GDE NORTE - COSERN', ' '],['20648', 'CEEE-D', 'CIA ESTADUAL DE DISTRIB ENER ELET-CEEE-D', 'N1'],['3204', 'CEEE-GT', 'CIA ESTADUAL GER.TRANS.ENER.ELET-CEEE-GT', 'N1'],['3069', 'FERBASA', 'CIA FERRO LIGAS DA BAHIA - FERBASA', 'N1'],['3077', 'CEDRO', 'CIA FIACAO TECIDOS CEDRO CACHOEIRA', 'N1'],['15636', 'COMGAS', 'CIA GAS DE SAO PAULO - COMGAS', ' '],['3298', 'HABITASUL', 'CIA HABITASUL DE PARTICIPACOES', ' '],['14761', 'CIA HERING', 'CIA HERING', 'NM'],['3395', 'IND CATAGUAS', 'CIA INDUSTRIAL CATAGUASES', ' '],['22691', 'LOCAMERICA', 'CIA LOCAÇÃO DAS AMÉRICAS', 'NM'],['3654', 'MELHOR SP', 'CIA MELHORAMENTOS DE SAO PAULO', ' '],['14311', 'COPEL', 'CIA PARANAENSE DE ENERGIA - COPEL', 'N1'],['18708', 'PAR AL BAHIA', 'CIA PARTICIPACOES ALIANCA DA BAHIA', ' '],['3824', 'PAUL F LUZ', 'CIA PAULISTA DE FORCA E LUZ', ' '],['19275', 'CPFL PIRATIN', 'CIA PIRATININGA DE FORCA E LUZ', ' '],['14443', 'SABESP', 'CIA SANEAMENTO BASICO EST SAO PAULO', 'NM'],['19445', 'COPASA', 'CIA SANEAMENTO DE MINAS GERAIS-COPASA MG', 'NM'],['18627', 'SANEPAR', 'CIA SANEAMENTO DO PARANA - SANEPAR', 'N2'],['3115', 'SEG AL BAHIA', 'CIA SEGUROS ALIANCA DA BAHIA', ' '],['4030', 'SID NACIONAL', 'CIA SIDERURGICA NACIONAL', ' '],['3158', 'COTEMINAS', 'CIA TECIDOS NORTE DE MINAS COTEMINAS', ' '],['4081', 'SANTANENSE', 'CIA TECIDOS SANTANENSE', ' '],['18287', 'CIBRASEC', 'CIBRASEC - COMPANHIA BRASILEIRA DE SECURITIZACAO', ' '],['21733', 'CIELO', 'CIELO S.A.', 'NM'],['14818', 'CIMS', 'CIMS S.A.', ' '],['23965', 'CINESYSTEM', 'CINESYSTEM S.A.', 'MA'],['17973', 'COGNA ON', 'COGNA EDUCAÇÃO S.A.', 'NM'],['22268', 'CONC RAPOSO', 'CONC AUTO RAPOSO TAVARES S.A.', ' '],['23515', 'GRUAIRPORT', 'CONC DO AEROPORTO INTERNACIONAL DE GUARULHOS S.A.', 'MB'],['20397', 'ECOVIAS', 'CONC ECOVIAS IMIGRANTES S.A.', ' '],['19208', 'CONC RIO TER', 'CONC RIO-TERESOPOLIS S.A.', 'MB'],['22411', 'ECOPISTAS', 'CONC ROD AYRTON SENNA E CARV PINTO S.A.-ECOPISTAS', ' '],['21024', 'VIAOESTE', 'CONC ROD.OESTE SP VIAOESTE S.A', ' '],['22721', 'ROD TIETE', 'CONC RODOVIAS DO TIETÊ S.A.', ' '],['22071', 'RT BANDEIRAS', 'CONC ROTA DAS BANDEIRAS S.A.', ' '],['20192', 'AUTOBAN', 'CONC SIST ANHANG-BANDEIRANT S.A. AUTOBAN', ' '],['4693', 'ODERICH', 'CONSERVAS ODERICH S.A.', ' '],['4707', 'ALFA CONSORC', 'CONSORCIO ALFA DE ADMINISTRACAO S.A.', ' '],['4723', 'CONST A LIND', 'CONSTRUTORA ADOLPHO LINDENBERG S.A.', ' '],['21148', 'TENDA', 'CONSTRUTORA TENDA S.A.', 'NM'],['4863', 'COR RIBEIRO', 'CORREA RIBEIRO S.A. COMERCIO E INDUSTRIA', ' '],['23485', 'COSAN LOG', 'COSAN LOGISTICA S.A.', 'NM'],['19836', 'COSAN', 'COSAN S.A.', 'NM'],['18660', 'CPFL ENERGIA', 'CPFL ENERGIA S.A.', 'NM'],['20540', 'CPFL RENOVAV', 'CPFL ENERGIAS RENOVÁVEIS S.A.', 'NM'],['18953', 'CPFL GERACAO', 'CPFL GERACAO DE ENERGIA S.A.', ' '],['20630', 'CR2', 'CR2 EMPREENDIMENTOS IMOBILIARIOS S.A.', 'NM'],['20044', 'CSU CARDSYST', 'CSU CARDSYSTEM S.A.', 'NM'],['23981', 'CTC S.A.', 'CTC - CENTRO DE TECNOLOGIA CANAVIEIRA S.A.', 'MA'],['18376', 'TRAN PAULIST', 'CTEEP - CIA TRANSMISSÃO ENERGIA ELÉTRICA PAULISTA', 'N1'],['23310', 'CVC BRASIL', 'CVC BRASIL OPERADORA E AGÊNCIA DE VIAGENS S.A.', 'NM'],['14460', 'CYRELA REALT', 'CYRELA BRAZIL REALTY S.A.EMPREEND E PART', 'NM'],['21040', 'CYRE COM-CCP', 'CYRELA COMMERCIAL PROPERT S.A. EMPR PART', 'NM'],['19623', 'DASA', 'DIAGNOSTICOS DA AMERICA S.A.', ' '],['14214', 'DIBENS LSG', 'DIBENS LEASING S.A. - ARREND.MERCANTIL', ' '],['9342', 'DIMED', 'DIMED S.A. DISTRIBUIDORA DE MEDICAMENTOS', ' '],['21350', 'DIRECIONAL', 'DIRECIONAL ENGENHARIA S.A.', 'NM'],['5207', 'DOHLER', 'DOHLER S.A.', ' '],['23493', 'DOMMO', 'DOMMO ENERGIA S.A.', ' '],['18597', 'DTCOM-DIRECT', 'DTCOM - DIRECT TO COMPANY S.A.', ' '],['21091', 'DURATEX', 'DURATEX S.A.', 'NM'],['21741', 'ECO SEC AGRO', 'ECO SECURITIZADORA DIREITOS CRED AGRONEGÓCIO S.A.', 'MB'],['21903', 'ECON', 'ECORODOVIAS CONCESSÕES E SERVIÇOS S.A.', ' '],['19453', 'ECORODOVIAS', 'ECORODOVIAS INFRAESTRUTURA E LOGÍSTICA S.A.', 'NM'],['19763', 'ENERGIAS BR', 'EDP - ENERGIAS DO BRASIL S.A.', 'NM'],['15342', 'ESCELSA', 'EDP ESPIRITO SANTO DISTRIBUIÇÃO DE ENERGIA S.A.', ' '],['16985', 'EBE', 'EDP SÃO PAULO DISTRIBUIÇÃO DE ENERGIA S.A.', ' '],['5380', 'ACO ALTONA', 'ELECTRO ACO ALTONA S.A.', ' '],['4359', 'ELEKEIROZ', 'ELEKEIROZ S.A.', ' '],['17485', 'ELEKTRO', 'ELEKTRO REDES S.A.', ' '],['15784', 'ELETROPAR', 'ELETROBRÁS PARTICIPAÇÕES S.A. - ELETROPAR', ' '],['14176', 'ELETROPAULO', 'ELETROPAULO METROP. ELET. SAO PAULO S.A.', ' '],['16993', 'EMAE', 'EMAE - EMPRESA METROP.AGUAS ENERGIA S.A.', ' '],['20087', 'EMBRAER', 'EMBRAER S.A.', 'NM'],['19011', 'ECONORTE', 'EMPRESA CONC RODOV DO NORTE S.A.ECONORTE', ' '],['16497', 'ENCORPAR', 'EMPRESA NAC COM REDITO PART S.A.ENCORPAR', ' '],['22365', 'ENAUTA PART', 'ENAUTA PARTICIPAÇÕES S.A.', 'NM'],['5576', 'ENERSUL', 'ENERGISA MATO GROSSO DO SUL - DIST DE ENERGIA S.A.', ' '],['14605', 'ENERGISA MT', 'ENERGISA MATO GROSSO-DISTRIBUIDORA DE ENERGIA S/A', ' '],['15253', 'ENERGISA', 'ENERGISA S.A.', 'N2'],['21237', 'ENEVA', 'ENEVA S.A', 'NM'],['17329', 'ENGIE BRASIL', 'ENGIE BRASIL ENERGIA S.A.', 'NM'],['20010', 'EQUATORIAL', 'EQUATORIAL ENERGIA S.A.', 'NM'],['16608', 'EQTLMARANHAO', 'EQUATORIAL MARANHÃO DISTRIBUIDORA DE ENERGIA S.A.', 'MB'],['18309', 'EQTL PARA', 'EQUATORIAL PARA DISTRIBUIDORA DE ENERGIA S.A.', ' '],['21016', 'YDUQS PART', 'ESTACIO PARTICIPACOES S.A.', 'NM'],['5762', 'ETERNIT', 'ETERNIT S.A.', 'NM'],['5770', 'EUCATEX', 'EUCATEX S.A. INDUSTRIA E COMERCIO', 'N1'],['20524', 'EVEN', 'EVEN CONSTRUTORA E INCORPORADORA S.A.', 'NM'],['1570', 'EXCELSIOR', 'EXCELSIOR ALIMENTOS S.A.', ' '],['20770', 'EZTEC', 'EZ TEC EMPREEND. E PARTICIPACOES S.A.', 'NM'],['22977', 'FGENERGIA', 'FERREIRA GOMES ENERGIA S.A.', ' '],['15369', 'FER C ATLANT', 'FERROVIA CENTRO-ATLANTICA S.A.', ' '],['20621', 'FER HERINGER', 'FERTILIZANTES HERINGER S.A.', 'NM'],['3891', 'ALFA FINANC', 'FINANCEIRA ALFA S.A.- CRED FINANC E INVS', ' '],['6076', 'FINANSINOS', 'FINANSINOS S.A.- CREDITO FINANC E INVEST', ' '],['21881', 'FLEURY', 'FLEURY S.A.', 'NM'],['24350', 'FLEX S/A', 'FLEX GESTÃO DE RELACIONAMENTOS S.A.', 'MA'],['6211', 'FRAS-LE', 'FRAS-LE S.A.', 'N1'],['16101', 'GAFISA', 'GAFISA S.A.', 'NM'],['22764', 'GAIA AGRO', 'GAIA AGRO SECURITIZADORA S.A.', ' '],['20222', 'GAIA SECURIT', 'GAIA SECURITIZADORA S.A.', 'MB'],['17965', 'GAMA PART', 'GAMA PARTICIPACOES S.A.', 'MB'],['21008', 'GENERALSHOPP', 'GENERAL SHOPPING E OUTLETS DO BRASIL S.A.', 'NM'],['3980', 'GERDAU', 'GERDAU S.A.', 'N1'],['19569', 'GOL', 'GOL LINHAS AEREAS INTELIGENTES S.A.', 'N2'],['80020', 'GP INVEST', 'GP INVESTMENTS. LTD.', 'DR3'],['16632', 'GPC PART', 'GPC PARTICIPACOES S.A.', ' '],['4537', 'GRAZZIOTIN', 'GRAZZIOTIN S.A.', ' '],['19615', 'GRENDENE', 'GRENDENE S.A.', 'NM'],['24694', 'CENTAURO', 'GRUPO SBF SA', 'NM'],['4669', 'GUARARAPES', 'GUARARAPES CONFECCOES S.A.', ' '],['13366', 'HAGA S/A', 'HAGA S.A. INDUSTRIA E COMERCIO', ' '],['24392', 'HAPVIDA', 'HAPVIDA PARTICIPACOES E INVESTIMENTOS SA', 'NM'],['20877', 'HELBOR', 'HELBOR EMPREENDIMENTOS S.A.', 'NM'],['6629', 'HERCULES', 'HERCULES S.A. FABRICA DE TALHERES', ' '],['6700', 'HOTEIS OTHON', 'HOTEIS OTHON S.A.', ' '],['21431', 'HYPERA', 'HYPERA S.A.', 'NM'],['18414', 'IDEIASNET', 'IDEIASNET S.A.', ' '],['6815', 'IGB S/A', 'IGB ELETRÔNICA S/A', ' '],['23175', 'IGUA SA', 'IGUA SANEAMENTO S.A.', 'MA'],['20494', 'IGUATEMI', 'IGUATEMI EMPRESA DE SHOPPING CENTERS S.A', 'NM'],['12319', 'J B DUARTE', 'INDUSTRIAS J B DUARTE S.A.', ' '],['7510', 'INDS ROMI', 'INDUSTRIAS ROMI S.A.', 'NM'],['7595', 'INEPAR', 'INEPAR S.A. INDUSTRIA E CONSTRUCOES', ' '],['17558', 'SELECTPART', 'INNCORP S.A.', 'MB'],['24090', 'IHPARDINI', 'INSTITUTO HERMES PARDINI S.A.', 'NM'],['24279', 'INTER SA', 'INTER CONSTRUTORA E INCORPORADORA S.A.', 'MA'],['23574', 'IMC S/A', 'INTERNATIONAL MEAL COMPANY ALIMENTACAO S.A.', 'NM'],['6041', 'INVEST BEMGE', 'INVESTIMENTOS BEMGE S.A.', ' '],['18775', 'INVEPAR', 'INVESTIMENTOS E PARTICIP. EM INFRA S.A. - INVEPAR', 'MB'],['11932', 'IOCHP-MAXION', 'IOCHPE MAXION S.A.', 'NM'],['2429', 'CELUL IRANI', 'IRANI PAPEL E EMBALAGEM S.A.', ' '],['24180', 'IRBBRASIL RE', 'IRB - BRASIL RESSEGUROS S.A.', 'NM'],['19364', 'ITAPEBI', 'ITAPEBI GERACAO DE ENERGIA S.A.', ' '],['19348', 'ITAUUNIBANCO', 'ITAU UNIBANCO HOLDING S.A.', 'N1'],['7617', 'ITAUSA', 'ITAUSA INVESTIMENTOS ITAU S.A.', 'N1'],['21156', 'J.MACEDO', 'J. MACEDO S.A.', ' '],['20575', 'JBS', 'JBS S.A.', 'NM'],['8672', 'JEREISSATI', 'JEREISSATI PARTICIPACOES S.A.', ' '],['20605', 'JHSF PART', 'JHSF PARTICIPACOES S.A.', 'NM'],['7811', 'JOAO FORTES', 'JOAO FORTES ENGENHARIA S.A.', ' '],['13285', 'JOSAPAR', 'JOSAPAR-JOAQUIM OLIVEIRA S.A. - PARTICIP', ' '],['22020', 'JSL', 'JSL S.A.', 'NM'],['4146', 'KARSTEN', 'KARSTEN S.A.', ' '],['7870', 'KEPLER WEBER', 'KEPLER WEBER S.A.', ' '],['12653', 'KLABIN S/A', 'KLABIN S.A.', 'N2'],['23434', 'LIBRA T RIO', 'LIBRA TERMINAL RIO S.A.', ' '],['24872', 'LIFEMED', 'LIFEMED INDUSTRIAL EQUIP. DE ART. MÉD. HOSP. S.A.', 'MA'],['19879', 'LIGHT S/A', 'LIGHT S.A.', 'NM'],['8036', 'LIGHT', 'LIGHT SERVICOS DE ELETRICIDADE S.A.', ' '],['23035', 'LINX', 'LINX S.A.', 'NM'],['19100', 'LIQ', 'LIQ PARTICIPAÇÕES S.A.', 'NM'],['15091', 'LITEL', 'LITEL PARTICIPACOES S.A.', 'MB'],['24759', 'LITELA', 'LITELA PARTICIPAÇÕES S.A.', 'MB'],['19739', 'LOCALIZA', 'LOCALIZA RENT A CAR S.A.', 'NM'],['24910', 'LOCAWEB', 'LOCAWEB SERVIÇOS DE INTERNET S.A.', 'NM'],['23272', 'LOG COM PROP', 'LOG COMMERCIAL PROPERTIES', 'NM'],['20710', 'LOG-IN', 'LOG-IN LOGISTICA INTERMODAL S.A.', 'NM'],['8087', 'LOJAS AMERIC', 'LOJAS AMERICANAS S.A.', 'N1'],['8133', 'LOJAS RENNER', 'LOJAS RENNER S.A.', 'NM'],['17434', 'LONGDIS', 'LONGDIS S.A.', 'MB'],['20370', 'LOPES BRASIL', 'LPS BRASIL - CONSULTORIA DE IMOVEIS S.A.', 'NM'],['20060', 'LUPATECH', 'LUPATECH S.A.', 'NM'],['20338', 'M.DIASBRANCO', 'M.DIAS BRANCO S.A. IND COM DE ALIMENTOS', 'NM'],['23612', 'MAESTROLOC', 'MAESTRO LOCADORA DE VEICULOS S.A.', 'MA'],['22470', 'MAGAZ LUIZA', 'MAGAZINE LUIZA S.A.', 'NM'],['8575', 'METAL LEVE', 'MAHLE-METAL LEVE S.A.', 'NM'],['8397', 'MANGELS INDL', 'MANGELS INDUSTRIAL S.A.', ' '],['8427', 'ESTRELA', 'MANUFATURA DE BRINQUEDOS ESTRELA S.A.', ' '],['8451', 'MARCOPOLO', 'MARCOPOLO S.A.', 'N2'],['20788', 'MARFRIG', 'MARFRIG GLOBAL FOODS S.A.', 'NM'],['22055', 'LOJAS MARISA', 'MARISA LOJAS S.A.', 'NM'],['8540', 'MERC FINANC', 'MERCANTIL BRASIL FINANC S.A. C.F.I.', ' '],['20613', 'METALFRIO', 'METALFRIO SOLUTIONS S.A.', 'NM'],['8605', 'METAL IGUACU', 'METALGRAFICA IGUACU S.A.', ' '],['8656', 'GERDAU MET', 'METALURGICA GERDAU S.A.', 'N1'],['13439', 'RIOSULENSE', 'METALURGICA RIOSULENSE S.A.', ' '],['8753', 'METISA', 'METISA METALURGICA TIMBOENSE S.A.', ' '],['22942', 'MGI PARTICIP', 'MGI - MINAS GERAIS PARTICIPAÇÕES S.A.', ' '],['22012', 'MILLS', 'MILLS ESTRUTURAS E SERVIÇOS DE ENGENHARIA S.A.', 'NM'],['8818', 'MINASMAQUINA', 'MINASMAQUINAS S.A.', ' '],['20931', 'MINERVA', 'MINERVA S.A.', 'NM'],['13765', 'MINUPAR', 'MINUPAR PARTICIPACOES S.A.', ' '],['24902', 'MITRE REALTY', 'MITRE REALTY EMPREENDIMENTOS E PARTICIPAÇÕES S.A.', 'NM'],['17914', 'MMX MINER', 'MMX MINERACAO E METALICOS S.A.', 'NM'],['8893', 'MONT ARANHA', 'MONTEIRO ARANHA S.A.', ' '],['21067', 'MOURA DUBEUX', 'MOURA DUBEUX ENGENHARIA S/A', 'NM'],['23825', 'MOVIDA', 'MOVIDA PARTICIPACOES SA', 'NM'],['17949', 'MRS LOGIST', 'MRS LOGISTICA S.A.', 'MB'],['20915', 'MRV', 'MRV ENGENHARIA E PARTICIPACOES S.A.', 'NM'],['20982', 'MULTIPLAN', 'MULTIPLAN - EMPREEND IMOBILIARIOS S.A.', 'N2'],['5312', 'MUNDIAL', 'MUNDIAL S.A. - PRODUTOS DE CONSUMO', ' '],['9040', 'NADIR FIGUEI', 'NADIR FIGUEIREDO IND E COM S.A.', ' '],['24783', 'GRUPO NATURA', 'NATURA &CO HOLDING S.A.', 'NM'],['19550', 'NATURA', 'NATURA COSMETICOS S.A.', ' '],['15539', 'NEOENERGIA', 'NEOENERGIA S.A.', 'NM'],['9083', 'NORDON MET', 'NORDON INDUSTRIAS METALURGICAS S.A.', ' '],['22985', 'NORTCQUIMICA', 'NORTEC QUÍMICA S.A.', 'MA'],['24384', 'INTERMEDICA', 'NOTRE DAME INTERMEDICA PARTICIPACOES SA', 'NM'],['21334', 'NUTRIPLANT', 'NUTRIPLANT INDUSTRIA E COMERCIO S.A.', 'MA'],['22390', 'OCTANTE SEC', 'OCTANTE SECURITIZADORA S.A.', ' '],['20125', 'ODONTOPREV', 'ODONTOPREV S.A.', 'NM'],['11312', 'OI', 'OI S.A.', 'N1'],['23426', 'OMEGA GER', 'OMEGA GERAÇÃO S.A.', 'NM'],['16942', 'OPPORT ENERG', 'OPPORTUNITY ENERGIA E PARTICIPACOES S.A.', 'MB'],['21342', 'OSX BRASIL', 'OSX BRASIL S.A.', 'NM'],['22250', 'OURINVESTSEC', 'OURINVEST SECURITIZADORA SA', ' '],['23507', 'OUROFINO S/A', 'OURO FINO SAUDE ANIMAL PARTICIPACOES S.A.', 'NM'],['23280', 'OURO VERDE', 'OURO VERDE LOCACAO E SERVICO S.A.', ' '],['94', 'PANATLANTICA', 'PANATLANTICA S.A.', ' '],['20729', 'PARANA', 'PARANA BCO S.A.', ' '],['9393', 'PARANAPANEMA', 'PARANAPANEMA S.A.', 'NM'],['18236', 'PATRIA SEC', 'PATRIA CIA SECURITIZADORA DE CRED IMOB', ' '],['13773', 'PORTOBELLO', 'PBG S/A', 'NM'],['21644', 'PDG SECURIT', 'PDG COMPANHIA SECURITIZADORA', ' '],['20478', 'PDG REALT', 'PDG REALTY S.A. EMPREEND E PARTICIPACOES', 'NM'],['22187', 'PETRORIO', 'PETRO RIO S.A.', 'NM'],['24295', 'PETROBRAS BR', 'PETROBRAS DISTRIBUIDORA S/A', 'NM'],['9512', 'PETROBRAS', 'PETROLEO BRASILEIRO S.A. PETROBRAS', 'N2'],['9539', 'PETTENATI', 'PETTENATI S.A. INDUSTRIA TEXTIL', ' '],['13471', 'PLASCAR PART', 'PLASCAR PARTICIPACOES INDUSTRIAIS S.A.', ' '],['22160', 'POLO CAP SEC', 'POLO CAPITAL SECURITIZADORA S.A', ' '],['13447', 'POLPAR', 'POLPAR S.A.', ' '],['19658', 'POMIFRUTAS', 'POMIFRUTAS S/A', 'NM'],['16659', 'PORTO SEGURO', 'PORTO SEGURO S.A.', 'NM'],['23523', 'PORTO VM', 'PORTO SUDESTE V.M. S.A.', ' '],['20362', 'POSITIVO TEC', 'POSITIVO TECNOLOGIA S.A.', 'NM'],['80152', 'PPLA', 'PPLA PARTICIPATIONS LTD.', 'DR3'],['24546', 'PRATICA', 'PRATICA KLIMAQUIP INDUSTRIA E COMERCIO SA', 'M2'],['24236', 'PRINER', 'PRINER SERVIÇOS INDUSTRIAIS S.A.', 'NM'],['19232', 'PROMAN', 'PRODUTORES ENERGET.DE MANSO S.A.- PROMAN', 'MB'],['20346', 'PROFARMA', 'PROFARMA DISTRIB PROD FARMACEUTICOS S.A.', 'NM'],['18333', 'PROMPT PART', 'PROMPT PARTICIPACOES S.A.', 'MB'],['22497', 'QUALICORP', 'QUALICORP CONSULTORIA E CORRETORA DE SEGUROS S.A.', 'NM'],['23302', 'QUALITY SOFT', 'QUALITY SOFTWARE S.A.', 'MA'],['5258', 'RAIADROGASIL', 'RAIA DROGASIL S.A.', 'NM'],['23230', 'RAIZEN ENERG', 'RAIZEN ENERGIA S.A.', ' '],['14109', 'RANDON PART', 'RANDON S.A. IMPLEMENTOS E PARTICIPACOES', 'N1'],['18406', 'RBCAPITALRES', 'RB CAPITAL COMPANHIA DE SECURITIZAÇÃO', 'MB'],['19860', 'RBCAPITALSEC', 'RB CAPITAL SECURITIZADORA S.A.', 'MB'],['18430', 'WTORRE PIC', 'REAL AI PIC SEC DE CREDITOS IMOBILIARIO S.A.', ' '],['12572', 'RECRUSUL', 'RECRUSUL S.A.', ' '],['3190', 'REDE ENERGIA', 'REDE ENERGIA PARTICIPAÇÕES S.A.', ' '],['9989', 'PET MANGUINH', 'REFINARIA DE PETROLEOS MANGUINHOS S.A.', ' '],['21636', 'RENOVA', 'RENOVA ENERGIA S.A.', 'N2'],['21440', 'LE LIS BLANC', 'RESTOQUE COMÉRCIO E CONFECÇÕES DE ROUPAS S.A.', 'NM'],['16527', 'AES SUL', 'RGE SUL DISTRIBUIDORA DE ENERGIA S.A.', ' '],['18368', 'GER PARANAP', 'RIO PARANAPANEMA ENERGIA S.A.', ' '],['20451', 'RNI', 'RNI NEGÓCIOS IMOBILIÁRIOS S.A.', 'NM'],['23167', 'ROD COLINAS', 'RODOVIAS DAS COLINAS S.A.', ' '],['16306', 'ROSSI RESID', 'ROSSI RESIDENCIAL S.A.', 'NM'],['15300', 'ALL NORTE', 'RUMO MALHA NORTE S.A.', 'MB'],['17930', 'ALL PAULISTA', 'RUMO MALHA PAULISTA S.A.', 'MB'],['17450', 'RUMO S.A.', 'RUMO S.A.', 'NM'],['23540', 'SALUS INFRA', 'SALUS INFRAESTRUTURA PORTUARIA SA', ' '],['19593', 'SANESALTO', 'SANESALTO SANEAMENTO S.A.', ' '],['12696', 'SANSUY', 'SANSUY S.A. INDUSTRIA DE PLASTICOS', ' '],['14923', 'SANTHER', 'SANTHER FAB DE PAPEL STA THEREZINHA S.A.', ' '],['23388', 'STO ANTONIO', 'SANTO ANTONIO ENERGIA S.A.', ' '],['17892', 'SANTOS BRP', 'SANTOS BRASIL PARTICIPACOES S.A.', 'NM'],['13781', 'SAO CARLOS', 'SAO CARLOS EMPREEND E PARTICIPACOES S.A.', 'NM'],['20516', 'SAO MARTINHO', 'SAO MARTINHO S.A.', 'NM'],['9415', 'SPTURIS', 'SAO PAULO TURISMO S.A.', ' '],['10472', 'SARAIVA LIVR', 'SARAIVA LIVREIROS S.A. - EM RECUPERAÇÃO JUDICIAL', 'N2'],['14664', 'SCHULZ', 'SCHULZ S.A.', ' '],['23221', 'SER EDUCA', 'SER EDUCACIONAL S.A.', 'NM'],['12823', 'ALIPERTI', 'SIDERURGICA J. L. ALIPERTI S.A.', ' '],['22799', 'SINQIA', 'SINQIA S.A.', 'NM'],['20745', 'SLC AGRICOLA', 'SLC AGRICOLA S.A.', 'NM'],['24260', 'SMART FIT', 'SMARTFIT ESCOLA DE GINÁSTICA E DANÇA S.A.', 'M2'],['24252', 'SMILES', 'SMILES FIDELIDADE S.A.', 'NM'],['10880', 'SONDOTECNICA', 'SONDOTECNICA ENGENHARIA SOLOS S.A.', ' '],['10960', 'SPRINGER', 'SPRINGER S.A.', ' '],['20966', 'SPRINGS', 'SPRINGS GLOBAL PARTICIPACOES S.A.', 'NM'],['24201', 'STARA', 'STARA S.A. - INDÚSTRIA DE IMPLEMENTOS AGRÍCOLAS', 'MA'],['22594', 'STATKRAFT', 'STATKRAFT ENERGIAS RENOVAVEIS S.A.', ' '],['16586', 'SUDESTE S/A', 'SUDESTE S.A.', 'MB'],['16438', 'SUL 116 PART', 'SUL 116 PARTICIPACOES S.A.', 'MB'],['21121', 'SUL AMERICA', 'SUL AMERICA S.A.', 'N2'],['9067', 'SUZANO HOLD', 'SUZANO HOLDING S.A.', ' '],['13986', 'SUZANO S.A.', 'SUZANO S.A.', 'NM'],['22454', 'TIME FOR FUN', 'T4F ENTRETENIMENTO S.A.', 'NM'],['6173', 'TAURUS ARMAS', 'TAURUS ARMAS S.A.', 'N2'],['24066', 'TCP TERMINAL', 'TCP TERMINAL DE CONTEINERES DE PARANAGUA SA', ' '],['22519', 'TECHNOS', 'TECHNOS S.A.', 'NM'],['20435', 'TECNISA', 'TECNISA S.A.', 'NM'],['11207', 'TECNOSOLO', 'TECNOSOLO ENGENHARIA S.A.', ' '],['20800', 'TEGMA', 'TEGMA GESTAO LOGISTICA S.A.', 'NM'],['11223', 'TEKA', 'TEKA-TECELAGEM KUEHNRICH S.A.', ' '],['11231', 'TEKNO', 'TEKNO S.A. - INDUSTRIA E COMERCIO', ' '],['11258', 'TELEBRAS', 'TELEC BRASILEIRAS S.A. TELEBRAS', ' '],['17671', 'TELEF BRASIL', 'TELEFÔNICA BRASIL S.A', ' '],['23329', 'TERM. PE III', 'TERMELÉTRICA PERNAMBUCO III S.A.', ' '],['18538', 'MENEZES CORT', 'TERMINAL GARAGEM MENEZES CORTES S.A.', 'MB'],['19852', 'TERMOPE', 'TERMOPERNAMBUCO S.A.', ' '],['20354', 'TERRA SANTA', 'TERRA SANTA AGRO S.A.', 'NM'],['7544', 'TEX RENAUX', 'TEXTIL RENAUXVIEW S.A.', ' '],['17639', 'TIM PART S/A', 'TIM PARTICIPACOES S.A.', 'NM'],['19992', 'TOTVS', 'TOTVS S.A.', 'NM'],['19330', 'TRIUNFO PART', 'TPI - TRIUNFO PARTICIP. E INVEST. S.A.', 'NM'],['20257', 'TAESA', 'TRANSMISSORA ALIANÇA DE ENERGIA ELÉTRICA S.A.', 'N2'],['8192', 'TREVISA', 'TREVISA INVESTIMENTOS S.A.', ' '],['23060', 'TRIANGULOSOL', 'TRIÂNGULO DO SOL AUTO-ESTRADAS S.A.', ' '],['21130', 'TRISUL', 'TRISUL S.A.', 'NM'],['11398', 'CRISTAL', 'TRONOX PIGMENTOS DO BRASIL S.A.', ' '],['22276', 'TRUESEC', 'TRUE SECURITIZADORA S.A.', ' '],['6343', 'TUPY', 'TUPY S.A.', 'NM'],['18465', 'ULTRAPAR', 'ULTRAPAR PARTICIPACOES S.A.', 'NM'],['22780', 'UNICASA', 'UNICASA INDÚSTRIA DE MÓVEIS S.A.', 'NM'],['21555', 'UNIDAS', 'UNIDAS S.A.', ' '],['11592', 'UNIPAR', 'UNIPAR CARBOCLORO S.A.', ' '],['16624', 'UPTICK', 'UPTICK PARTICIPACOES S.A.', 'MB'],['14320', 'USIMINAS', 'USINAS SID DE MINAS GERAIS S.A.-USIMINAS', 'N1'],['4170', 'VALE', 'VALE S.A.', 'NM'],['20028', 'VALID', 'VALID SOLUÇÕES S.A.', 'NM'],['23990', 'VERTCIASEC', 'VERT COMPANHIA SECURITIZADORA', ' '],['6505', 'VIAVAREJO', 'VIA VAREJO S.A.', 'NM'],['24805', 'VIVARA S.A.', 'VIVARA PARTICIPAÇOES S.A', 'NM'],['20702', 'VIVER', 'VIVER INCORPORADORA E CONSTRUTORA S.A.', 'NM'],['11762', 'VULCABRAS', 'VULCABRAS/AZALEIA S.A.', 'NM'],['5410', 'WEG', 'WEG S.A.', 'NM'],['11991', 'WETZEL S/A', 'WETZEL S.A.', ' '],['14346', 'WHIRLPOOL', 'WHIRLPOOL S.A.', ' '],['80047', 'WILSON SONS', 'WILSON SONS LTD.', 'DR3'],['23590', 'WIZ S.A.', 'WIZ SOLUÇÕES E CORRETAGEM DE SEGUROS S.A.', 'NM'],['11070', 'WLM IND COM', 'WLM PART. E COMÉRCIO DE MÁQUINAS E VEÍCULOS S.A.', ' ']]
        print('...done')
    except Exception as e:
        restart(e, __name__)
def getSheetOfCompaniesDEBUG():
    try:
        print('LOAD sheet of companies')

        print('0 Companies updated to sheet')
        print('...done')
    except Exception as e:
        restart(e, 'updatedListOfCompanies')
def sortListOfCompaniesByCompanyDEBUG():
    try:
        global list_of_companies
        list_of_companies = list_of_companies = [['5410', 'WEG', 'WEG S.A.', 'NM'],['11991', 'WETZEL S/A', 'WETZEL S.A.', ' '],['14346', 'WHIRLPOOL', 'WHIRLPOOL S.A.', ' '],['80047', 'WILSON SONS', 'WILSON SONS LTD.', 'DR3'],['23590', 'WIZ S.A.', 'WIZ SOLUÇÕES E CORRETAGEM DE SEGUROS S.A.', 'NM'],['11070', 'WLM IND COM', 'WLM PART. E COMÉRCIO DE MÁQUINAS E VEÍCULOS S.A.', ' '],['16284', '524 PARTICIP', '524 PARTICIPACOES S.A.', 'MB'],['20958', 'ABC BRASIL', 'BCO ABC BRASIL S.A.', 'N2'],['5380', 'ACO ALTONA', 'ELECTRO ACO ALTONA S.A.', ' '],['21725', 'ADVANCED-DH', 'ADVANCED DIGITAL HEALTH MEDICINA PREVENTIVA S.A.', ' '],['16527', 'AES SUL', 'RGE SUL DISTRIBUIDORA DE ENERGIA S.A.', ' '],['18970', 'AES TIETE E', 'AES TIETE ENERGIA SA', 'N2'],['22179', 'AFLUENTE T', 'AFLUENTE TRANSMISSÃO DE ENERGIA ELÉTRICA S/A', ' '],['16705', 'ALEF S/A', 'ALEF S.A.', 'MB'],['4707', 'ALFA CONSORC', 'CONSORCIO ALFA DE ADMINISTRACAO S.A.', ' '],['3891', 'ALFA FINANC', 'FINANCEIRA ALFA S.A.- CRED FINANC E INVS', ' '],['9954', 'ALFA HOLDING', 'ALFA HOLDINGS S.A.', ' '],['1384', 'ALFA INVEST', 'BCO ALFA DE INVESTIMENTO S.A.', ' '],['21032', 'ALGAR TELEC', 'ALGAR TELECOM S/A', ' '],['22357', 'ALIANSCSONAE', 'ALIANSCE SONAE SHOPPING CENTERS S.A.', 'NM'],['12823', 'ALIPERTI', 'SIDERURGICA J. L. ALIPERTI S.A.', ' '],['15300', 'ALL NORTE', 'RUMO MALHA NORTE S.A.', 'MB'],['17930', 'ALL PAULISTA', 'RUMO MALHA PAULISTA S.A.', 'MB'],['24058', 'ALLIAR', 'CENTRO DE IMAGEM DIAGNOSTICOS S.A.', 'NM'],['10456', 'ALPARGATAS', 'ALPARGATAS S.A.', 'N1'],['22217', 'ALPER S.A.', 'ALPER CONSULTORIA E CORRETORA DE SEGUROS S.A.', 'NM'],['18066', 'ALTERE SEC', 'ALTERE SECURITIZADORA S.A.', ' '],['21490', 'ALUPAR', 'ALUPAR INVESTIMENTO S/A', 'N2'],['922', 'AMAZONIA', 'BCO AMAZONIA S.A.', ' '],['23264', 'AMBEV S/A', 'AMBEV S.A.', ' '],['3050', 'AMPLA ENERG', 'AMPLA ENERGIA E SERVICOS S.A.', ' '],['23248', 'ANIMA', 'ANIMA HOLDING S.A.', 'NM'],['22349', 'AREZZO CO', 'AREZZO INDÚSTRIA E COMÉRCIO S.A.', 'NM'],['15423', 'ATOMPAR', 'ATOM EMPREENDIMENTOS E PARTICIPAÇÕES S.A.', ' '],['20192', 'AUTOBAN', 'CONC SIST ANHANG-BANDEIRANT S.A. AUTOBAN', ' '],['11975', 'AZEVEDO', 'AZEVEDO E TRAVASSOS S.A.', ' '],['24112', 'AZUL', 'AZUL S.A.', 'N2'],['20990', 'B2W DIGITAL', 'B2W - COMPANHIA DIGITAL', 'NM'],['21610', 'B3', 'B3 S.A. - BRASIL. BOLSA. BALCÃO', 'NM'],['701', 'BAHEMA', 'BAHEMA EDUCAÇÃO S.A.', 'MA'],['24600', 'BANCO BMG', 'BANCO BMG S.A.', 'N1'],['24406', 'BANCO INTER', 'BANCO INTER S.A.', 'N2'],['21199', 'BANCO PAN', 'BCO PAN S.A.', 'N1'],['1120', 'BANESE', 'BCO ESTADO DE SERGIPE S.A. - BANESE', ' '],['1155', 'BANESTES', 'BANESTES S.A. - BCO EST ESPIRITO SANTO', ' '],['1171', 'BANPARA', 'BCO ESTADO DO PARA S.A.', ' '],['1210', 'BANRISUL', 'BCO ESTADO DO RIO GRANDE DO SUL S.A.', 'N1'],['1520', 'BARDELLA', 'BARDELLA S.A. INDUSTRIAS MECANICAS', ' '],['15458', 'BATTISTELLA', 'BATTISTELLA ADM PARTICIPACOES S.A.', ' '],['1562', 'BAUMER', 'BAUMER S.A.', ' '],['24660', 'BBMLOGISTICA', 'BBM LOGISTICA S.A.', 'MA'],['23159', 'BBSEGURIDADE', 'BB SEGURIDADE PARTICIPAÇÕES S.A.', 'NM'],['19747', 'BETA SECURIT', 'BETA SECURITIZADORA S.A.', ' '],['17884', 'BETAPART', 'BETAPART PARTICIPACOES S.A.', 'MB'],['1694', 'BIC MONARK', 'BICICLETAS MONARK S.A.', ' '],['19305', 'BIOMM', 'BIOMM S.A.', 'MA'],['22845', 'BIOSEV', 'BIOSEV S.A.', 'NM'],['80179', 'BIOTOSCANA', 'BIOTOSCANA INVESTMENTS S.A.', 'DR3'],['24317', 'BK BRASIL', 'BK BRASIL OPERAÇÃO E ASSESSORIA A RESTAURANTES SA', 'NM'],['16772', 'BNDESPAR', 'BNDES PARTICIPACOES S.A. - BNDESPAR', 'MB'],['12190', 'BOMBRIL', 'BOMBRIL S.A.', ' '],['21180', 'BR BROKERS', 'BRASIL BROKERS PARTICIPACOES S.A.', 'NM'],['19909', 'BR MALLS PAR', 'BR MALLS PARTICIPACOES S.A.', 'NM'],['19925', 'BR PROPERT', 'BR PROPERTIES S.A.', 'NM'],['906', 'BRADESCO', 'BCO BRADESCO S.A.', 'N1'],['19640', 'BRADESCO LSG', 'BRADESCO LEASING S.A. ARREND MERCANTIL', ' '],['18724', 'BRADESPAR', 'BRADESPAR S.A.', 'N1'],['1023', 'BRASIL', 'BCO BRASIL S.A.', 'NM'],['20036', 'BRASILAGRO', 'BRASILAGRO - CIA BRAS DE PROP AGRICOLAS', 'NM'],['4820', 'BRASKEM', 'BRASKEM S.A.', 'N1'],['19720', 'BRAZIL REALT', 'BRAZIL REALTY CIA SECURIT. CRÉD. IMOBILIÁRIOS', ' '],['17922', 'BRAZILIAN FR', 'BRAZILIAN FINANCE E REAL ESTATE S.A.', ' '],['18759', 'BRAZILIAN SC', 'BRAZILIAN SECURITIES CIA SECURITIZACAO', ' '],['14206', 'BRB BANCO', 'BRB BCO DE BRASILIA S.A.', ' '],['20672', 'BRC SECURIT', 'BRC SECURITIZADORA S.A.', ' '],['16292', 'BRF SA', 'BRF S.A.', 'NM'],['19984', 'BRPR 56 SEC', 'BRPR 56 SECURITIZADORA CRED IMOB S.A.', ' '],['23817', 'BRQ', 'BRQ SOLUCOES EM INFORMATICA S.A.', 'MA'],['22616', 'BTGP BANCO', 'BCO BTG PACTUAL S.A.', 'N2'],['20133', 'BV LEASING', 'BV LEASING - ARRENDAMENTO MERCANTIL S.A.', ' '],['19119', 'CABINDA PART', 'CABINDA PARTICIPACOES S.A.', 'MB'],['22683', 'CACHOEIRA', 'CACHOEIRA PAULISTA TRANSMISSORA ENERGIA S.A.', 'MB'],['19135', 'CACONDE PART', 'CACONDE PARTICIPACOES S.A.', 'MB'],['2100', 'CAMBUCI', 'CAMBUCI S.A.', ' '],['24228', 'CAMIL', 'CAMIL ALIMENTOS S.A.', 'NM'],['17493', 'CAPITALPART', 'CAPITALPART PARTICIPACOES S.A.', 'MB'],['24171', 'CARREFOUR BR', 'ATACADÃO S.A.', 'NM'],['16861', 'CASAN', 'CIA CATARINENSE DE AGUAS E SANEAM.-CASAN', ' '],['18821', 'CCR SA', 'CCR S.A.', 'NM'],['24848', 'CEA MODAS', 'CEA MODAS S.A.', 'NM'],['14451', 'CEB', 'CIA ENERGETICA DE BRASILIA', ' '],['3077', 'CEDRO', 'CIA FIACAO TECIDOS CEDRO CACHOEIRA', 'N1'],['20648', 'CEEE-D', 'CIA ESTADUAL DE DISTRIB ENER ELET-CEEE-D', 'N1'],['3204', 'CEEE-GT', 'CIA ESTADUAL GER.TRANS.ENER.ELET-CEEE-GT', 'N1'],['16616', 'CEG', 'CIA DISTRIB DE GAS DO RIO DE JANEIRO-CEG', ' '],['2461', 'CELESC', 'CENTRAIS ELET DE SANTA CATARINA S.A.', 'N2'],['21393', 'CELGPAR', 'CIA CELG DE PARTICIPACOES - CELGPAR', ' '],['14362', 'CELPE', 'CIA ENERGETICA DE PERNAMBUCO - CELPE', ' '],['2429', 'CELUL IRANI', 'IRANI PAPEL E EMBALAGEM S.A.', ' '],['13854', 'CEMEPE', 'CEMEPE INVESTIMENTOS S.A.', ' '],['2453', 'CEMIG', 'CIA ENERGETICA DE MINAS GERAIS - CEMIG', 'N1'],['20303', 'CEMIG DIST', 'CEMIG DISTRIBUICAO S.A.', ' '],['20320', 'CEMIG GT', 'CEMIG GERACAO E TRANSMISSAO S.A.', ' '],['24694', 'CENTAURO', 'GRUPO SBF SA', 'NM'],['2577', 'CESP', 'CESP - CIA ENERGETICA DE SAO PAULO', 'N1'],['14761', 'CIA HERING', 'CIA HERING', 'NM'],['18287', 'CIBRASEC', 'CIBRASEC - COMPANHIA BRASILEIRA DE SECURITIZACAO', ' '],['21733', 'CIELO', 'CIELO S.A.', 'NM'],['14818', 'CIMS', 'CIMS S.A.', ' '],['23965', 'CINESYSTEM', 'CINESYSTEM S.A.', 'MA'],['14524', 'COELBA', 'CIA ELETRICIDADE EST. DA BAHIA - COELBA', ' '],['14869', 'COELCE', 'CIA ENERGETICA DO CEARA - COELCE', ' '],['17973', 'COGNA ON', 'COGNA EDUCAÇÃO S.A.', 'NM'],['15636', 'COMGAS', 'CIA GAS DE SAO PAULO - COMGAS', ' '],['22268', 'CONC RAPOSO', 'CONC AUTO RAPOSO TAVARES S.A.', ' '],['19208', 'CONC RIO TER', 'CONC RIO-TERESOPOLIS S.A.', 'MB'],['4723', 'CONST A LIND', 'CONSTRUTORA ADOLPHO LINDENBERG S.A.', ' '],['19445', 'COPASA', 'CIA SANEAMENTO DE MINAS GERAIS-COPASA MG', 'NM'],['14311', 'COPEL', 'CIA PARANAENSE DE ENERGIA - COPEL', 'N1'],['4863', 'COR RIBEIRO', 'CORREA RIBEIRO S.A. COMERCIO E INDUSTRIA', ' '],['19836', 'COSAN', 'COSAN S.A.', 'NM'],['23485', 'COSAN LOG', 'COSAN LOGISTICA S.A.', 'NM'],['18139', 'COSERN', 'CIA ENERGETICA DO RIO GDE NORTE - COSERN', ' '],['3158', 'COTEMINAS', 'CIA TECIDOS NORTE DE MINAS COTEMINAS', ' '],['18660', 'CPFL ENERGIA', 'CPFL ENERGIA S.A.', 'NM'],['18953', 'CPFL GERACAO', 'CPFL GERACAO DE ENERGIA S.A.', ' '],['19275', 'CPFL PIRATIN', 'CIA PIRATININGA DE FORCA E LUZ', ' '],['20540', 'CPFL RENOVAV', 'CPFL ENERGIAS RENOVÁVEIS S.A.', 'NM'],['20630', 'CR2', 'CR2 EMPREENDIMENTOS IMOBILIARIOS S.A.', 'NM'],['11398', 'CRISTAL', 'TRONOX PIGMENTOS DO BRASIL S.A.', ' '],['20044', 'CSU CARDSYST', 'CSU CARDSYSTEM S.A.', 'NM'],['23981', 'CTC S.A.', 'CTC - CENTRO DE TECNOLOGIA CANAVIEIRA S.A.', 'MA'],['23310', 'CVC BRASIL', 'CVC BRASIL OPERADORA E AGÊNCIA DE VIAGENS S.A.', 'NM'],['21040', 'CYRE COM-CCP', 'CYRELA COMMERCIAL PROPERT S.A. EMPR PART', 'NM'],['14460', 'CYRELA REALT', 'CYRELA BRAZIL REALTY S.A.EMPREEND E PART', 'NM'],['19623', 'DASA', 'DIAGNOSTICOS DA AMERICA S.A.', ' '],['14214', 'DIBENS LSG', 'DIBENS LEASING S.A. - ARREND.MERCANTIL', ' '],['9342', 'DIMED', 'DIMED S.A. DISTRIBUIDORA DE MEDICAMENTOS', ' '],['21350', 'DIRECIONAL', 'DIRECIONAL ENGENHARIA S.A.', 'NM'],['5207', 'DOHLER', 'DOHLER S.A.', ' '],['23493', 'DOMMO', 'DOMMO ENERGIA S.A.', ' '],['18597', 'DTCOM-DIRECT', 'DTCOM - DIRECT TO COMPANY S.A.', ' '],['21091', 'DURATEX', 'DURATEX S.A.', 'NM'],['16985', 'EBE', 'EDP SÃO PAULO DISTRIBUIÇÃO DE ENERGIA S.A.', ' '],['21741', 'ECO SEC AGRO', 'ECO SECURITIZADORA DIREITOS CRED AGRONEGÓCIO S.A.', 'MB'],['21903', 'ECON', 'ECORODOVIAS CONCESSÕES E SERVIÇOS S.A.', ' '],['19011', 'ECONORTE', 'EMPRESA CONC RODOV DO NORTE S.A.ECONORTE', ' '],['22411', 'ECOPISTAS', 'CONC ROD AYRTON SENNA E CARV PINTO S.A.-ECOPISTAS', ' '],['19453', 'ECORODOVIAS', 'ECORODOVIAS INFRAESTRUTURA E LOGÍSTICA S.A.', 'NM'],['20397', 'ECOVIAS', 'CONC ECOVIAS IMIGRANTES S.A.', ' '],['4359', 'ELEKEIROZ', 'ELEKEIROZ S.A.', ' '],['17485', 'ELEKTRO', 'ELEKTRO REDES S.A.', ' '],['2437', 'ELETROBRAS', 'CENTRAIS ELET BRAS S.A. - ELETROBRAS', 'N1'],['15784', 'ELETROPAR', 'ELETROBRÁS PARTICIPAÇÕES S.A. - ELETROPAR', ' '],['14176', 'ELETROPAULO', 'ELETROPAULO METROP. ELET. SAO PAULO S.A.', ' '],['16993', 'EMAE', 'EMAE - EMPRESA METROP.AGUAS ENERGIA S.A.', ' '],['20087', 'EMBRAER', 'EMBRAER S.A.', 'NM'],['22365', 'ENAUTA PART', 'ENAUTA PARTICIPAÇÕES S.A.', 'NM'],['16497', 'ENCORPAR', 'EMPRESA NAC COM REDITO PART S.A.ENCORPAR', ' '],['19763', 'ENERGIAS BR', 'EDP - ENERGIAS DO BRASIL S.A.', 'NM'],['15253', 'ENERGISA', 'ENERGISA S.A.', 'N2'],['14605', 'ENERGISA MT', 'ENERGISA MATO GROSSO-DISTRIBUIDORA DE ENERGIA S/A', ' '],['5576', 'ENERSUL', 'ENERGISA MATO GROSSO DO SUL - DIST DE ENERGIA S.A.', ' '],['21237', 'ENEVA', 'ENEVA S.A', 'NM'],['17329', 'ENGIE BRASIL', 'ENGIE BRASIL ENERGIA S.A.', 'NM'],['18309', 'EQTL PARA', 'EQUATORIAL PARA DISTRIBUIDORA DE ENERGIA S.A.', ' '],['16608', 'EQTLMARANHAO', 'EQUATORIAL MARANHÃO DISTRIBUIDORA DE ENERGIA S.A.', 'MB'],['20010', 'EQUATORIAL', 'EQUATORIAL ENERGIA S.A.', 'NM'],['15342', 'ESCELSA', 'EDP ESPIRITO SANTO DISTRIBUIÇÃO DE ENERGIA S.A.', ' '],['8427', 'ESTRELA', 'MANUFATURA DE BRINQUEDOS ESTRELA S.A.', ' '],['5762', 'ETERNIT', 'ETERNIT S.A.', 'NM'],['5770', 'EUCATEX', 'EUCATEX S.A. INDUSTRIA E COMERCIO', 'N1'],['20524', 'EVEN', 'EVEN CONSTRUTORA E INCORPORADORA S.A.', 'NM'],['1570', 'EXCELSIOR', 'EXCELSIOR ALIMENTOS S.A.', ' '],['20770', 'EZTEC', 'EZ TEC EMPREEND. E PARTICIPACOES S.A.', 'NM'],['15369', 'FER C ATLANT', 'FERROVIA CENTRO-ATLANTICA S.A.', ' '],['20621', 'FER HERINGER', 'FERTILIZANTES HERINGER S.A.', 'NM'],['3069', 'FERBASA', 'CIA FERRO LIGAS DA BAHIA - FERBASA', 'N1'],['22977', 'FGENERGIA', 'FERREIRA GOMES ENERGIA S.A.', ' '],['6076', 'FINANSINOS', 'FINANSINOS S.A.- CREDITO FINANC E INVEST', ' '],['21881', 'FLEURY', 'FLEURY S.A.', 'NM'],['24350', 'FLEX S/A', 'FLEX GESTÃO DE RELACIONAMENTOS S.A.', 'MA'],['6211', 'FRAS-LE', 'FRAS-LE S.A.', 'N1'],['16101', 'GAFISA', 'GAFISA S.A.', 'NM'],['22764', 'GAIA AGRO', 'GAIA AGRO SECURITIZADORA S.A.', ' '],['20222', 'GAIA SECURIT', 'GAIA SECURITIZADORA S.A.', 'MB'],['17965', 'GAMA PART', 'GAMA PARTICIPACOES S.A.', 'MB'],['21008', 'GENERALSHOPP', 'GENERAL SHOPPING E OUTLETS DO BRASIL S.A.', 'NM'],['18368', 'GER PARANAP', 'RIO PARANAPANEMA ENERGIA S.A.', ' '],['3980', 'GERDAU', 'GERDAU S.A.', 'N1'],['8656', 'GERDAU MET', 'METALURGICA GERDAU S.A.', 'N1'],['19569', 'GOL', 'GOL LINHAS AEREAS INTELIGENTES S.A.', 'N2'],['80020', 'GP INVEST', 'GP INVESTMENTS. LTD.', 'DR3'],['16632', 'GPC PART', 'GPC PARTICIPACOES S.A.', ' '],['4537', 'GRAZZIOTIN', 'GRAZZIOTIN S.A.', ' '],['19615', 'GRENDENE', 'GRENDENE S.A.', 'NM'],['23515', 'GRUAIRPORT', 'CONC DO AEROPORTO INTERNACIONAL DE GUARULHOS S.A.', 'MB'],['24783', 'GRUPO NATURA', 'NATURA &CO HOLDING S.A.', 'NM'],['4669', 'GUARARAPES', 'GUARARAPES CONFECCOES S.A.', ' '],['3298', 'HABITASUL', 'CIA HABITASUL DE PARTICIPACOES', ' '],['13366', 'HAGA S/A', 'HAGA S.A. INDUSTRIA E COMERCIO', ' '],['24392', 'HAPVIDA', 'HAPVIDA PARTICIPACOES E INVESTIMENTOS SA', 'NM'],['20877', 'HELBOR', 'HELBOR EMPREENDIMENTOS S.A.', 'NM'],['6629', 'HERCULES', 'HERCULES S.A. FABRICA DE TALHERES', ' '],['6700', 'HOTEIS OTHON', 'HOTEIS OTHON S.A.', ' '],['21431', 'HYPERA', 'HYPERA S.A.', 'NM'],['18414', 'IDEIASNET', 'IDEIASNET S.A.', ' '],['6815', 'IGB S/A', 'IGB ELETRÔNICA S/A', ' '],['23175', 'IGUA SA', 'IGUA SANEAMENTO S.A.', 'MA'],['20494', 'IGUATEMI', 'IGUATEMI EMPRESA DE SHOPPING CENTERS S.A', 'NM'],['24090', 'IHPARDINI', 'INSTITUTO HERMES PARDINI S.A.', 'NM'],['23574', 'IMC S/A', 'INTERNATIONAL MEAL COMPANY ALIMENTACAO S.A.', 'NM'],['3395', 'IND CATAGUAS', 'CIA INDUSTRIAL CATAGUASES', ' '],['7510', 'INDS ROMI', 'INDUSTRIAS ROMI S.A.', 'NM'],['20885', 'INDUSVAL', 'BCO INDUSVAL S.A.', 'N2'],['7595', 'INEPAR', 'INEPAR S.A. INDUSTRIA E CONSTRUCOES', ' '],['24279', 'INTER SA', 'INTER CONSTRUTORA E INCORPORADORA S.A.', 'MA'],['24384', 'INTERMEDICA', 'NOTRE DAME INTERMEDICA PARTICIPACOES SA', 'NM'],['18775', 'INVEPAR', 'INVESTIMENTOS E PARTICIP. EM INFRA S.A. - INVEPAR', 'MB'],['6041', 'INVEST BEMGE', 'INVESTIMENTOS BEMGE S.A.', ' '],['11932', 'IOCHP-MAXION', 'IOCHPE MAXION S.A.', 'NM'],['24180', 'IRBBRASIL RE', 'IRB - BRASIL RESSEGUROS S.A.', 'NM'],['19364', 'ITAPEBI', 'ITAPEBI GERACAO DE ENERGIA S.A.', ' '],['7617', 'ITAUSA', 'ITAUSA INVESTIMENTOS ITAU S.A.', 'N1'],['19348', 'ITAUUNIBANCO', 'ITAU UNIBANCO HOLDING S.A.', 'N1'],['12319', 'J B DUARTE', 'INDUSTRIAS J B DUARTE S.A.', ' '],['21156', 'J.MACEDO', 'J. MACEDO S.A.', ' '],['20575', 'JBS', 'JBS S.A.', 'NM'],['8672', 'JEREISSATI', 'JEREISSATI PARTICIPACOES S.A.', ' '],['20605', 'JHSF PART', 'JHSF PARTICIPACOES S.A.', 'NM'],['7811', 'JOAO FORTES', 'JOAO FORTES ENGENHARIA S.A.', ' '],['13285', 'JOSAPAR', 'JOSAPAR-JOAQUIM OLIVEIRA S.A. - PARTICIP', ' '],['22020', 'JSL', 'JSL S.A.', 'NM'],['4146', 'KARSTEN', 'KARSTEN S.A.', ' '],['7870', 'KEPLER WEBER', 'KEPLER WEBER S.A.', ' '],['12653', 'KLABIN S/A', 'KLABIN S.A.', 'N2'],['21440', 'LE LIS BLANC', 'RESTOQUE COMÉRCIO E CONFECÇÕES DE ROUPAS S.A.', 'NM'],['23434', 'LIBRA T RIO', 'LIBRA TERMINAL RIO S.A.', ' '],['24872', 'LIFEMED', 'LIFEMED INDUSTRIAL EQUIP. DE ART. MÉD. HOSP. S.A.', 'MA'],['8036', 'LIGHT', 'LIGHT SERVICOS DE ELETRICIDADE S.A.', ' '],['19879', 'LIGHT S/A', 'LIGHT S.A.', 'NM'],['23035', 'LINX', 'LINX S.A.', 'NM'],['19100', 'LIQ', 'LIQ PARTICIPAÇÕES S.A.', 'NM'],['15091', 'LITEL', 'LITEL PARTICIPACOES S.A.', 'MB'],['24759', 'LITELA', 'LITELA PARTICIPAÇÕES S.A.', 'MB'],['19739', 'LOCALIZA', 'LOCALIZA RENT A CAR S.A.', 'NM'],['22691', 'LOCAMERICA', 'CIA LOCAÇÃO DAS AMÉRICAS', 'NM'],['24910', 'LOCAWEB', 'LOCAWEB SERVIÇOS DE INTERNET S.A.', 'NM'],['23272', 'LOG COM PROP', 'LOG COMMERCIAL PROPERTIES', 'NM'],['20710', 'LOG-IN', 'LOG-IN LOGISTICA INTERMODAL S.A.', 'NM'],['8087', 'LOJAS AMERIC', 'LOJAS AMERICANAS S.A.', 'N1'],['22055', 'LOJAS MARISA', 'MARISA LOJAS S.A.', 'NM'],['8133', 'LOJAS RENNER', 'LOJAS RENNER S.A.', 'NM'],['17434', 'LONGDIS', 'LONGDIS S.A.', 'MB'],['20370', 'LOPES BRASIL', 'LPS BRASIL - CONSULTORIA DE IMOVEIS S.A.', 'NM'],['20060', 'LUPATECH', 'LUPATECH S.A.', 'NM'],['20338', 'M.DIASBRANCO', 'M.DIAS BRANCO S.A. IND COM DE ALIMENTOS', 'NM'],['23612', 'MAESTROLOC', 'MAESTRO LOCADORA DE VEICULOS S.A.', 'MA'],['22470', 'MAGAZ LUIZA', 'MAGAZINE LUIZA S.A.', 'NM'],['8397', 'MANGELS INDL', 'MANGELS INDUSTRIAL S.A.', ' '],['8451', 'MARCOPOLO', 'MARCOPOLO S.A.', 'N2'],['20788', 'MARFRIG', 'MARFRIG GLOBAL FOODS S.A.', 'NM'],['3654', 'MELHOR SP', 'CIA MELHORAMENTOS DE SAO PAULO', ' '],['18538', 'MENEZES CORT', 'TERMINAL GARAGEM MENEZES CORTES S.A.', 'MB'],['1325', 'MERC BRASIL', 'BCO MERCANTIL DO BRASIL S.A.', ' '],['8540', 'MERC FINANC', 'MERCANTIL BRASIL FINANC S.A. C.F.I.', ' '],['1309', 'MERC INVEST', 'BCO MERCANTIL DE INVESTIMENTOS S.A.', ' '],['8605', 'METAL IGUACU', 'METALGRAFICA IGUACU S.A.', ' '],['8575', 'METAL LEVE', 'MAHLE-METAL LEVE S.A.', 'NM'],['20613', 'METALFRIO', 'METALFRIO SOLUTIONS S.A.', 'NM'],['8753', 'METISA', 'METISA METALURGICA TIMBOENSE S.A.', ' '],['22942', 'MGI PARTICIP', 'MGI - MINAS GERAIS PARTICIPAÇÕES S.A.', ' '],['22012', 'MILLS', 'MILLS ESTRUTURAS E SERVIÇOS DE ENGENHARIA S.A.', 'NM'],['8818', 'MINASMAQUINA', 'MINASMAQUINAS S.A.', ' '],['20931', 'MINERVA', 'MINERVA S.A.', 'NM'],['13765', 'MINUPAR', 'MINUPAR PARTICIPACOES S.A.', ' '],['24902', 'MITRE REALTY', 'MITRE REALTY EMPREENDIMENTOS E PARTICIPAÇÕES S.A.', 'NM'],['17914', 'MMX MINER', 'MMX MINERACAO E METALICOS S.A.', 'NM'],['8893', 'MONT ARANHA', 'MONTEIRO ARANHA S.A.', ' '],['21067', 'MOURA DUBEUX', 'MOURA DUBEUX ENGENHARIA S/A', 'NM'],['23825', 'MOVIDA', 'MOVIDA PARTICIPACOES SA', 'NM'],['17949', 'MRS LOGIST', 'MRS LOGISTICA S.A.', 'MB'],['20915', 'MRV', 'MRV ENGENHARIA E PARTICIPACOES S.A.', 'NM'],['20982', 'MULTIPLAN', 'MULTIPLAN - EMPREEND IMOBILIARIOS S.A.', 'N2'],['5312', 'MUNDIAL', 'MUNDIAL S.A. - PRODUTOS DE CONSUMO', ' '],['9040', 'NADIR FIGUEI', 'NADIR FIGUEIREDO IND E COM S.A.', ' '],['19550', 'NATURA', 'NATURA COSMETICOS S.A.', ' '],['15539', 'NEOENERGIA', 'NEOENERGIA S.A.', 'NM'],['1228', 'NORD BRASIL', 'BCO NORDESTE DO BRASIL S.A.', ' '],['9083', 'NORDON MET', 'NORDON INDUSTRIAS METALURGICAS S.A.', ' '],['22985', 'NORTCQUIMICA', 'NORTEC QUÍMICA S.A.', 'MA'],['21334', 'NUTRIPLANT', 'NUTRIPLANT INDUSTRIA E COMERCIO S.A.', 'MA'],['22390', 'OCTANTE SEC', 'OCTANTE SECURITIZADORA S.A.', ' '],['4693', 'ODERICH', 'CONSERVAS ODERICH S.A.', ' '],['20125', 'ODONTOPREV', 'ODONTOPREV S.A.', 'NM'],['11312', 'OI', 'OI S.A.', 'N1'],['23426', 'OMEGA GER', 'OMEGA GERAÇÃO S.A.', 'NM'],['16942', 'OPPORT ENERG', 'OPPORTUNITY ENERGIA E PARTICIPACOES S.A.', 'MB'],['21342', 'OSX BRASIL', 'OSX BRASIL S.A.', 'NM'],['22250', 'OURINVESTSEC', 'OURINVEST SECURITIZADORA SA', ' '],['23280', 'OURO VERDE', 'OURO VERDE LOCACAO E SERVICO S.A.', ' '],['23507', 'OUROFINO S/A', 'OURO FINO SAUDE ANIMAL PARTICIPACOES S.A.', 'NM'],['14826', 'P.ACUCAR-CBD', 'CIA BRASILEIRA DE DISTRIBUICAO', 'N1'],['94', 'PANATLANTICA', 'PANATLANTICA S.A.', ' '],['18708', 'PAR AL BAHIA', 'CIA PARTICIPACOES ALIANCA DA BAHIA', ' '],['20729', 'PARANA', 'PARANA BCO S.A.', ' '],['9393', 'PARANAPANEMA', 'PARANAPANEMA S.A.', 'NM'],['18236', 'PATRIA SEC', 'PATRIA CIA SECURITIZADORA DE CRED IMOB', ' '],['3824', 'PAUL F LUZ', 'CIA PAULISTA DE FORCA E LUZ', ' '],['20478', 'PDG REALT', 'PDG REALTY S.A. EMPREEND E PARTICIPACOES', 'NM'],['21644', 'PDG SECURIT', 'PDG COMPANHIA SECURITIZADORA', ' '],['9989', 'PET MANGUINH', 'REFINARIA DE PETROLEOS MANGUINHOS S.A.', ' '],['9512', 'PETROBRAS', 'PETROLEO BRASILEIRO S.A. PETROBRAS', 'N2'],['24295', 'PETROBRAS BR', 'PETROBRAS DISTRIBUIDORA S/A', 'NM'],['22187', 'PETRORIO', 'PETRO RIO S.A.', 'NM'],['9539', 'PETTENATI', 'PETTENATI S.A. INDUSTRIA TEXTIL', ' '],['20567', 'PINE', 'BCO PINE S.A.', 'N2'],['13471', 'PLASCAR PART', 'PLASCAR PARTICIPACOES INDUSTRIAIS S.A.', ' '],['22160', 'POLO CAP SEC', 'POLO CAPITAL SECURITIZADORA S.A', ' '],['13447', 'POLPAR', 'POLPAR S.A.', ' '],['19658', 'POMIFRUTAS', 'POMIFRUTAS S/A', 'NM'],['16659', 'PORTO SEGURO', 'PORTO SEGURO S.A.', 'NM'],['23523', 'PORTO VM', 'PORTO SUDESTE V.M. S.A.', ' '],['13773', 'PORTOBELLO', 'PBG S/A', 'NM'],['20362', 'POSITIVO TEC', 'POSITIVO TECNOLOGIA S.A.', 'NM'],['80152', 'PPLA', 'PPLA PARTICIPATIONS LTD.', 'DR3'],['24546', 'PRATICA', 'PRATICA KLIMAQUIP INDUSTRIA E COMERCIO SA', 'M2'],['24236', 'PRINER', 'PRINER SERVIÇOS INDUSTRIAIS S.A.', 'NM'],['20346', 'PROFARMA', 'PROFARMA DISTRIB PROD FARMACEUTICOS S.A.', 'NM'],['19232', 'PROMAN', 'PRODUTORES ENERGET.DE MANSO S.A.- PROMAN', 'MB'],['18333', 'PROMPT PART', 'PROMPT PARTICIPACOES S.A.', 'MB'],['22497', 'QUALICORP', 'QUALICORP CONSULTORIA E CORRETORA DE SEGUROS S.A.', 'NM'],['23302', 'QUALITY SOFT', 'QUALITY SOFTWARE S.A.', 'MA'],['5258', 'RAIADROGASIL', 'RAIA DROGASIL S.A.', 'NM'],['23230', 'RAIZEN ENERG', 'RAIZEN ENERGIA S.A.', ' '],['14109', 'RANDON PART', 'RANDON S.A. IMPLEMENTOS E PARTICIPACOES', 'N1'],['18406', 'RBCAPITALRES', 'RB CAPITAL COMPANHIA DE SECURITIZAÇÃO', 'MB'],['19860', 'RBCAPITALSEC', 'RB CAPITAL SECURITIZADORA S.A.', 'MB'],['12572', 'RECRUSUL', 'RECRUSUL S.A.', ' '],['3190', 'REDE ENERGIA', 'REDE ENERGIA PARTICIPAÇÕES S.A.', ' '],['21636', 'RENOVA', 'RENOVA ENERGIA S.A.', 'N2'],['13439', 'RIOSULENSE', 'METALURGICA RIOSULENSE S.A.', ' '],['20451', 'RNI', 'RNI NEGÓCIOS IMOBILIÁRIOS S.A.', 'NM'],['23167', 'ROD COLINAS', 'RODOVIAS DAS COLINAS S.A.', ' '],['22721', 'ROD TIETE', 'CONC RODOVIAS DO TIETÊ S.A.', ' '],['16306', 'ROSSI RESID', 'ROSSI RESIDENCIAL S.A.', 'NM'],['22071', 'RT BANDEIRAS', 'CONC ROTA DAS BANDEIRAS S.A.', ' '],['17450', 'RUMO S.A.', 'RUMO S.A.', 'NM'],['14443', 'SABESP', 'CIA SANEAMENTO BASICO EST SAO PAULO', 'NM'],['23540', 'SALUS INFRA', 'SALUS INFRAESTRUTURA PORTUARIA SA', ' '],['18627', 'SANEPAR', 'CIA SANEAMENTO DO PARANA - SANEPAR', 'N2'],['19593', 'SANESALTO', 'SANESALTO SANEAMENTO S.A.', ' '],['12696', 'SANSUY', 'SANSUY S.A. INDUSTRIA DE PLASTICOS', ' '],['20532', 'SANTANDER BR', 'BCO SANTANDER (BRASIL) S.A.', ' '],['4081', 'SANTANENSE', 'CIA TECIDOS SANTANENSE', ' '],['14923', 'SANTHER', 'SANTHER FAB DE PAPEL STA THEREZINHA S.A.', ' '],['17892', 'SANTOS BRP', 'SANTOS BRASIL PARTICIPACOES S.A.', 'NM'],['13781', 'SAO CARLOS', 'SAO CARLOS EMPREEND E PARTICIPACOES S.A.', 'NM'],['20516', 'SAO MARTINHO', 'SAO MARTINHO S.A.', 'NM'],['10472', 'SARAIVA LIVR', 'SARAIVA LIVREIROS S.A. - EM RECUPERAÇÃO JUDICIAL', 'N2'],['14664', 'SCHULZ', 'SCHULZ S.A.', ' '],['3115', 'SEG AL BAHIA', 'CIA SEGUROS ALIANCA DA BAHIA', ' '],['17558', 'SELECTPART', 'INNCORP S.A.', 'MB'],['23221', 'SER EDUCA', 'SER EDUCACIONAL S.A.', 'NM'],['4030', 'SID NACIONAL', 'CIA SIDERURGICA NACIONAL', ' '],['22799', 'SINQIA', 'SINQIA S.A.', 'NM'],['20745', 'SLC AGRICOLA', 'SLC AGRICOLA S.A.', 'NM'],['24260', 'SMART FIT', 'SMARTFIT ESCOLA DE GINÁSTICA E DANÇA S.A.', 'M2'],['24252', 'SMILES', 'SMILES FIDELIDADE S.A.', 'NM'],['10880', 'SONDOTECNICA', 'SONDOTECNICA ENGENHARIA SOLOS S.A.', ' '],['10960', 'SPRINGER', 'SPRINGER S.A.', ' '],['20966', 'SPRINGS', 'SPRINGS GLOBAL PARTICIPACOES S.A.', 'NM'],['9415', 'SPTURIS', 'SAO PAULO TURISMO S.A.', ' '],['24201', 'STARA', 'STARA S.A. - INDÚSTRIA DE IMPLEMENTOS AGRÍCOLAS', 'MA'],['22594', 'STATKRAFT', 'STATKRAFT ENERGIAS RENOVAVEIS S.A.', ' '],['23388', 'STO ANTONIO', 'SANTO ANTONIO ENERGIA S.A.', ' '],['16586', 'SUDESTE S/A', 'SUDESTE S.A.', 'MB'],['16438', 'SUL 116 PART', 'SUL 116 PARTICIPACOES S.A.', 'MB'],['21121', 'SUL AMERICA', 'SUL AMERICA S.A.', 'N2'],['9067', 'SUZANO HOLD', 'SUZANO HOLDING S.A.', ' '],['13986', 'SUZANO S.A.', 'SUZANO S.A.', 'NM'],['20257', 'TAESA', 'TRANSMISSORA ALIANÇA DE ENERGIA ELÉTRICA S.A.', 'N2'],['6173', 'TAURUS ARMAS', 'TAURUS ARMAS S.A.', 'N2'],['24066', 'TCP TERMINAL', 'TCP TERMINAL DE CONTEINERES DE PARANAGUA SA', ' '],['22519', 'TECHNOS', 'TECHNOS S.A.', 'NM'],['20435', 'TECNISA', 'TECNISA S.A.', 'NM'],['11207', 'TECNOSOLO', 'TECNOSOLO ENGENHARIA S.A.', ' '],['20800', 'TEGMA', 'TEGMA GESTAO LOGISTICA S.A.', 'NM'],['11223', 'TEKA', 'TEKA-TECELAGEM KUEHNRICH S.A.', ' '],['11231', 'TEKNO', 'TEKNO S.A. - INDUSTRIA E COMERCIO', ' '],['11258', 'TELEBRAS', 'TELEC BRASILEIRAS S.A. TELEBRAS', ' '],['17671', 'TELEF BRASIL', 'TELEFÔNICA BRASIL S.A', ' '],['21148', 'TENDA', 'CONSTRUTORA TENDA S.A.', 'NM'],['23329', 'TERM. PE III', 'TERMELÉTRICA PERNAMBUCO III S.A.', ' '],['19852', 'TERMOPE', 'TERMOPERNAMBUCO S.A.', ' '],['20354', 'TERRA SANTA', 'TERRA SANTA AGRO S.A.', 'NM'],['7544', 'TEX RENAUX', 'TEXTIL RENAUXVIEW S.A.', ' '],['17639', 'TIM PART S/A', 'TIM PARTICIPACOES S.A.', 'NM'],['22454', 'TIME FOR FUN', 'T4F ENTRETENIMENTO S.A.', 'NM'],['19992', 'TOTVS', 'TOTVS S.A.', 'NM'],['18376', 'TRAN PAULIST', 'CTEEP - CIA TRANSMISSÃO ENERGIA ELÉTRICA PAULISTA', 'N1'],['8192', 'TREVISA', 'TREVISA INVESTIMENTOS S.A.', ' '],['23060', 'TRIANGULOSOL', 'TRIÂNGULO DO SOL AUTO-ESTRADAS S.A.', ' '],['21130', 'TRISUL', 'TRISUL S.A.', 'NM'],['19330', 'TRIUNFO PART', 'TPI - TRIUNFO PARTICIP. E INVEST. S.A.', 'NM'],['22276', 'TRUESEC', 'TRUE SECURITIZADORA S.A.', ' '],['6343', 'TUPY', 'TUPY S.A.', 'NM'],['18465', 'ULTRAPAR', 'ULTRAPAR PARTICIPACOES S.A.', 'NM'],['22780', 'UNICASA', 'UNICASA INDÚSTRIA DE MÓVEIS S.A.', 'NM'],['21555', 'UNIDAS', 'UNIDAS S.A.', ' '],['11592', 'UNIPAR', 'UNIPAR CARBOCLORO S.A.', ' '],['16624', 'UPTICK', 'UPTICK PARTICIPACOES S.A.', 'MB'],['14320', 'USIMINAS', 'USINAS SID DE MINAS GERAIS S.A.-USIMINAS', 'N1'],['4170', 'VALE', 'VALE S.A.', 'NM'],['20028', 'VALID', 'VALID SOLUÇÕES S.A.', 'NM'],['23990', 'VERTCIASEC', 'VERT COMPANHIA SECURITIZADORA', ' '],['21024', 'VIAOESTE', 'CONC ROD.OESTE SP VIAOESTE S.A', ' '],['6505', 'VIAVAREJO', 'VIA VAREJO S.A.', 'NM'],['24805', 'VIVARA S.A.', 'VIVARA PARTICIPAÇOES S.A', 'NM'],['20702', 'VIVER', 'VIVER INCORPORADORA E CONSTRUTORA S.A.', 'NM'],['11762', 'VULCABRAS', 'VULCABRAS/AZALEIA S.A.', 'NM'],['18430', 'WTORRE PIC', 'REAL AI PIC SEC DE CREDITOS IMOBILIARIO S.A.', ' '],['21016', 'YDUQS PART', 'ESTACIO PARTICIPACOES S.A.', 'NM']]
        return list_of_companies
        print('...done')
    except Exception as e:
        restart(e, __name__)
def getCompanyDEBUG(company):
    try:
        print('GATHERING COMPANY INFO - ', company[1], 'DEBUG')

        getCompanyMainPageDEBUG(company)
        getCompanyListOfReportsDEBUG(company)

        print('...done')
    except Exception as e:
        restart(e, __name__)
def getCompanyMainPageDEBUG(c):
    try:
        global company

        company = c
        print('... Company Details DEBUG')
        company.extend(['27/02/2020 23:31:00',['WEGE3'], '84.429.695/0001-11',
                   'A Weg Sa É Uma Sociedade de Participação Não Operacional (holding) E também Sociedade de Comando Do Grupo Weg.',
                   'Bens Industriais', 'Máquinas e Equipamentos', 'Motores . Compressores e Outros', 'www.weg.net/br',
                   'on skipped', 'pn skipped'])
    except Exception as e:
        restart(e, __name__)
def getCompanyListOfReportsDEBUG(c):
    try:
        global company

        print('... Company Reports DEBUG')
        company_reports = [['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=91065&CodigoTipoInstituicao=2', '91065', '20191231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=88758&CodigoTipoInstituicao=2', '88758', '20190930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=85540&CodigoTipoInstituicao=2', '85540', '20190630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=82335&CodigoTipoInstituicao=2', '82335', '20190331', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=82128&CodigoTipoInstituicao=2', '82128', '20181231', 'Demonstrações Financeiras Padronizadas', 'Versão 2.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=78246&CodigoTipoInstituicao=2', '78246', '20180930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=77198&CodigoTipoInstituicao=2', '77198', '20180630', 'Informações Trimestrais', 'Versão 2.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=78619&CodigoTipoInstituicao=2', '78619', '20180331', 'Informações Trimestrais', 'Versão 2.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=71853&CodigoTipoInstituicao=2', '71853', '20171231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=69250&CodigoTipoInstituicao=2', '69250', '20170930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=67211&CodigoTipoInstituicao=2', '67211', '20170630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=64436&CodigoTipoInstituicao=2', '64436', '20170331', 'Informações Trimestrais', 'Versão 2.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=62688&CodigoTipoInstituicao=2', '62688', '20161231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=60179&CodigoTipoInstituicao=2', '60179', '20160930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=58280&CodigoTipoInstituicao=2', '58280', '20160630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=55186&CodigoTipoInstituicao=2', '55186', '20160331', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=53579&CodigoTipoInstituicao=2', '53579', '20151231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=51132&CodigoTipoInstituicao=2', '51132', '20150930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=49293&CodigoTipoInstituicao=2', '49293', '20150630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=46193&CodigoTipoInstituicao=2', '46193', '20150331', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=44241&CodigoTipoInstituicao=2', '44241', '20141231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=41847&CodigoTipoInstituicao=2', '41847', '20140930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=39852&CodigoTipoInstituicao=2', '39852', '20140630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=36378&CodigoTipoInstituicao=2', '36378', '20140331', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=34835&CodigoTipoInstituicao=2', '34835', '20131231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=32134&CodigoTipoInstituicao=2', '32134', '20130930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=30081&CodigoTipoInstituicao=2', '30081', '20130630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=26269&CodigoTipoInstituicao=2', '26269', '20130331', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=24482&CodigoTipoInstituicao=2', '24482', '20121231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=21857&CodigoTipoInstituicao=2', '21857', '20120930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=19852&CodigoTipoInstituicao=2', '19852', '20120630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=16424&CodigoTipoInstituicao=2', '16424', '20120331', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=14329&CodigoTipoInstituicao=2', '14329', '20111231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=12156&CodigoTipoInstituicao=2', '12156', '20110930', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=10245&CodigoTipoInstituicao=2', '10245', '20110630', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=6694&CodigoTipoInstituicao=2', '6694', '20110331', 'Informações Trimestrais', 'Versão 1.0'],['http://www.rad.cvm.gov.br/ENETCONSULTA/frmGerenciaPaginaFRE.aspx?NumeroSequencialDocumento=4906&CodigoTipoInstituicao=2', '4906', '20101231', 'Demonstrações Financeiras Padronizadas', 'Versão 1.0'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2010&tipo=4', '5410', '20100930', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2010&tipo=4', '5410', '20100630', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2010&tipo=4', '5410', '20100331', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2009&tipo=2', '5410', '20091231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2009&tipo=4', '5410', '20090930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2009&tipo=4', '5410', '20090630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2009&tipo=4', '5410', '20090331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2008&tipo=2', '5410', '20081231', 'Demonstrações Financeiras Padronizadas', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2008&tipo=4', '5410', '20080930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2008&tipo=4', '5410', '20080630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2008&tipo=4', '5410', '20080331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2007&tipo=2', '5410', '20071231', 'Demonstrações Financeiras Padronizadas', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2007&tipo=4', '5410', '20070930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2007&tipo=4', '5410', '20070630', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2007&tipo=4', '5410', '20070331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2006&tipo=2', '5410', '20061231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2006&tipo=4', '5410', '20060930', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2006&tipo=4', '5410', '20060630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2006&tipo=4', '5410', '20060331', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2005&tipo=2', '5410', '20051231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2005&tipo=4', '5410', '20050930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2005&tipo=4', '5410', '20050630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2005&tipo=4', '5410', '20050331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2004&tipo=2', '5410', '20041231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2004&tipo=4', '5410', '20040930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2004&tipo=4', '5410', '20040630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2004&tipo=4', '5410', '20040331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2003&tipo=2', '5410', '20031231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2003&tipo=4', '5410', '20030930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2003&tipo=4', '5410', '20030630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2003&tipo=4', '5410', '20030331', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2002&tipo=2', '5410', '20021231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2002&tipo=4', '5410', '20020930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2002&tipo=4', '5410', '20020630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2002&tipo=4', '5410', '20020331', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2001&tipo=2', '5410', '20011231', 'Demonstrações Financeiras Padronizadas', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2001&tipo=4', '5410', '20010930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2001&tipo=4', '5410', '20010630', 'Informações Trimestrais', 'Reapresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2001&tipo=4', '5410', '20010331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/2000&tipo=2', '5410', '20001231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/2000&tipo=4', '5410', '20000930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/2000&tipo=4', '5410', '20000630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/2000&tipo=4', '5410', '20000331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/1999&tipo=2', '5410', '19991231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/1999&tipo=4', '5410', '19990930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/06/1999&tipo=4', '5410', '19990630', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/03/1999&tipo=4', '5410', '19990331', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/1998&tipo=2', '5410', '19981231', 'Demonstrações Financeiras Padronizadas', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=30/09/1998&tipo=4', '5410', '19980930', 'Informações Trimestrais', 'Apresentação'],['http://www2.bmfbovespa.com.br/dxw/FrDXW.asp?site=B&mercado=18&razao=WEG S.A.&pregao=WEG&ccvm=5410&data=31/12/1997&tipo=2', '5410', '19971231', 'Demonstrações Financeiras Padronizadas', 'Apresentação']]
        company.append(company_reports)
    except Exception as e:
        restart(e, __name__)
def updateCompanyToSheetDEBUG(company):
    try:
        if google != True:
            googleAPI()

        # create or update sheet
        global newsheet_id

        print('UPDATE INFO to Google Sheets DEBUG')

        print('UPDATED worksheets with new info for', company[1], company[5][0][:4].replace('Nenh','NONE'))

    except Exception as e:
        restart(e, __name__)
def vacationLogDEBUG(c):
    try:
        if google != True:
            googleAPI()
        company = c
        newsheet_id = '1BzAzPbnu9HJ23xVi7IkOi24549XeeAOaHbHUq-psPsk'

        print('LOG Company', company[1])

        # update log sheet
        log_data = [str(company[5][0][:4].replace('Nenh','NONE')), 'https://docs.google.com/spreadsheets/d/' + newsheet_id, timestamp]
        ws_bovespa_log.append_row(log_data)

        # cell = ws_bovespa_lista_bovespa.find(company[1])
        # row = str(cell).split(' ')[1][1:].split('C')[0]
        # ws_bovespa_lista_bovespa.update_cell(row, 5, datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        # print('...done')

        # data = ws_bovespa_lista_bovespa.get_all_values()
        # search = company[1]
        # for row, sublist in enumerate(data):
        #     if sublist[1] == search:
        #         ws_bovespa_lista_bovespa.update_cell(row+1, 5, company[4])
        #         break

        print('...done')
    except Exception as e:
        restart(e, __name__)


# System Warp-Up/Cool Down Block
def start():
    try:
        global google
        google = False
        global browser
        global wait
        global timestamp

        # start engines
        print('-- Hey Ho,')
        # load browser and general parameters
        browser = webdriver.Chrome(executable_path="C:/Users/faust/PycharmProjects/chromedriver.exe")
        # browser = webdriver.Firefox(executable_path="C:/Users/faust/Documents/python drivers/geckodriver-v0.26.0-win64/geckodriver.exe")
        wait = WebDriverWait(browser, 60)
        browser.minimize_window()
        timestamp = datetime.now()
        timestamp = timestamp.strftime("%d/%m/%Y %H:%M:%S")
        print(timestamp)
        print('Let\'s Go!')
    except Exception as e:
        restart(e, __name__)
def end():
    try:
        print('-- This is the end, my friend')
        browser.quit()
        quit()
    except Exception as e:
        print('-- Ops, Sheet Happens!')
        quit()
def restart(e, msg):
    try:
        # original
        # def restart(e, msg):
        #     browser.quit()
        #     print('stop working in', msg, e)
        #     eastern_egg_project()
        #     # quit()

        browser.quit()
        print('stop working in', msg, e)
        # summer_project()
        # vacation_project()
        winter_project()
        # spring_project()
        # quit()
    except:
        print('Erro terminal desconhecido. A coisa foi grave!')
        browser.quit()
        # eastern_egg_project()
        vacation_project()
        quit()


# Projects Block
def summer_project():
    try:
        print('SUMMER PROJECT in action')
        # run_in__logical_steps
        a = user_defined_variables()
        b = start()

        # c = getListOfCompaniesDEBUG()
        c = getListOfCompanies()

        # d = getSheetOfCompaniesDEBUG()
        d = getSheetOfCompanies()

        # e = sortListOfCompaniesByCompanyDEBUG()
        e = sortListOfCompaniesByCompany()

    except Exception as e:
        restart(e, __name__)
def vacation_project():
    try:
        print('VACATION PROJECT in action')
        # run_in__logical_steps
        a = user_defined_variables()
        b = start()

        c = getListOfCompaniesDEBUG()
        # c = getListOfCompanies()

        # d = getSheetOfCompaniesDEBUG()
        d = getSheetOfCompanies()

        # e = sortListOfCompaniesByCompanyDEBUG()
        e = sortListOfCompaniesByCompany()

        for i, company in enumerate(list_of_companies):
            if i < batch_companies:
                # f = getCompanyDEBUG(company)
                f = getCompany(company)

                # g = updateCompanyToSheetDEBUG(company)
                g = updateCompanyToSheet(company)

                # z = vacationLogDEBUG(company)
                z = vacationLog(company, 5)

    except Exception as e:
        restart(e, __name__)
def winter_project():
    try:
        print('WINTER PROJECT in action')
        # run_in__logical_steps
        a = user_defined_variables()
        b = start()

        c = getListOfCompaniesDEBUG()
        # c = getListOfCompanies()

        # d = getSheetOfCompaniesDEBUG()
        d = getSheetOfCompanies()

        # e = sortListOfCompaniesByCompanyDEBUG()
        e = sortListOfCompaniesByReport()

        for i, company in enumerate(list_of_companies):
            if i < batch_companies:
                # f = getCompanyDEBUG(company)
                # f = getCompany(company)

                # g = updateCompanyToSheetDEBUG(company)
                # g = updateCompanyToSheet(company)

                h = getSheetListOfReports(company)

                # z = vacationLogDEBUG(company)
                z = vacationLog(company, 6)

    except Exception as e:
        restart(e, __name__)
def spring_project():
    try:
        print('SPRING 4 PROJECT in action')
        # run_in__logical_steps
        a = user_defined_variables()
        b = start()

        c = getListOfCompaniesDEBUG()
        # c = getListOfCompanies()

        # d = getSheetOfCompaniesDEBUG()
        d = getSheetOfCompanies()

        # e = sortListOfCompaniesByCompanyDEBUG()
        e = sortListOfCompaniesByReport()

        f = updateUberblasterlista()

    except Exception as e:
        restart(e, __name__)
0

# run_in_a_line_geek
# aa = summer_project()  # list of companies
# bb = vacation_project()  # company sheet and data
cc = winter_project()  # reports content
# dd = spring_project() # datastudio
z = end()

