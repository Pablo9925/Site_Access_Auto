#!/usr/bin/env python
# coding: utf-8

# In[1]:


import re
import csv
import time
import pyautogui
from database import areas, sitios, solicitantes
from selenium import webdriver
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from selenium.webdriver.support.ui import Select

user = 'exjcar90'
password = 'K2kOctub2023+'
beginning = 2
end = 100
#API sheets
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
KEY = 'key.json'
SPREADSHEET_ID = '1to5Q1qMQ63FhqWADiKfHiT4MFtVq4NeKp0u0yLid7-o'
creds = service_account.Credentials.from_service_account_file(KEY, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)
sheet = service.spreadsheets()
file_path_minutograma = r'C:\Users\Usuario\Desktop\K2K\Site access\Minutograma - Varias Actividades.xlsx'
file_path_poliza = r'C:\Users\Usuario\Desktop\K2K\Site access\K2K Soluciones S.A.S._Póliza Contrato 40364 Instalación equipos 2G, 3G, 4G, 5G (1).pdf'

def get_sheet_values(sheet, column, beginning, end):
    result = sheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f'App!{column}{beginning}:{column}{end}'
    ).execute()
    return result.get('values', [])

def get_spreadsheet_data(sheet, beginning, end):
    solicitante = get_sheet_values(sheet, "B", beginning, end)
    approver = get_sheet_values(sheet, "C", beginning, end)
    tel_approver = get_sheet_values(sheet, "D", beginning, end)
    lider = get_sheet_values(sheet, "E", beginning, end)
    site1 = get_sheet_values(sheet, "F", beginning, end)
    site2 = get_sheet_values(sheet, "G", beginning, end)
    site3 = get_sheet_values(sheet, "H", beginning, end)
    site4 = get_sheet_values(sheet, "I", beginning, end)
    site5 = get_sheet_values(sheet, "J", beginning, end)
    comentarios = get_sheet_values(sheet, "L", beginning, end)
    elementos = get_sheet_values(sheet, "M", beginning, end)
    actividad = get_sheet_values(sheet, "N", beginning, end)
    ID = get_sheet_values(sheet, "K", beginning, end)
    return solicitante, approver, tel_approver, lider, site1, site2, site3, site4, site5, comentarios, elementos, actividad, ID

def append_values_to_spreadsheet(sheet, spreadsheet_id, range_, values, sheet_name):
    result = sheet.values().append(
        spreadsheetId=spreadsheet_id,
        range=f'{sheet_name}!{range_}',
        valueInputOption='USER_ENTERED',
        body={'values': values}
    ).execute()
    return result

def login(driver,user,password):
    #Pasar restricción de seguridad
    details = driver.find_element("id", "details-button")
    details.click()
    details = driver.find_element("id", "proceed-link")
    details.click()
    #login
    username_field = driver.find_element(By.ID, 'user_app')
    password_field = driver.find_element(By.ID, 'passwd_app')
    username_field.send_keys(user)
    password_field.send_keys(password)
    pyautogui.press('enter')

def find_and_send_keys(id,key):
    time.sleep(0.5)
    element = wait.until(EC.presence_of_element_located((By.ID, id)))
    element.send_keys(key)

def find_and_click(id):
    time.sleep(0.5)
    element = wait.until(EC.presence_of_element_located((By.ID, id)))
    element.click()

def find_and_click_upload():
    time.sleep(0.5)
    upload_icon = driver.find_element(By.CSS_SELECTOR, "i.fas.fa-upload")
    upload_icon.click()
    
def find_and_click_by_xpath(xpath):
    time.sleep(0.5)
    element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    element.click()
    
def find_and_click_by_value(id, value):
    time.sleep(0.5)
    xpath = f"//input[@id='{id}' and @value='{value}']"
    element = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
    element.click()

def find_element(id,text):
    time.sleep(0.5)
    element = wait.until(EC.presence_of_element_located((By.ID, id)))
    element.click()
    select = Select(element)
    select.select_by_visible_text(text)
    
def prints(id):
    time.sleep(0.5)
    element = wait.until(EC.presence_of_element_located((By.ID, id)))
    element.click()
    select = Select(element)
    for option in select.options:
        print(option.text)
        
def find_and_send_date(id, key):
    time.sleep(0.5)
    element = wait.until(EC.presence_of_element_located((By.ID, id)))
    driver.execute_script("arguments[0].removeAttribute('readonly')", element)
    element.clear()
    element.send_keys(key)

SOLICITANTE, APPROVER, TEL_APPROVER, LIDER, SITE1, SITE2, SITE3, SITE4, SITE5, COMENTARIOS, ELEMENTOS, ACTIVIDAD, ID = get_spreadsheet_data(sheet, beginning, end)

#webdriver
driver = webdriver.Chrome()
driver.get('https://10.67.106.100/app/login.php')

login(driver,user,password)

wait = WebDriverWait(driver, 20)  # Se define una espera de hasta 20 segundos
wait.until(EC.url_changes('https://10.67.106.100/app/login.php'))

for i in range(len(SOLICITANTE)):
    if i>=len(ID):
        site = []
        solicitante = SOLICITANTE[i][0]
        approver = APPROVER[i][0]
        tel_approver = TEL_APPROVER[i][0]
        try:
            site.append(SITE1[i][0])
        except IndexError:
            pass
        try:
            site.append(SITE2[i][0])
        except IndexError:
            pass
        try:
            site.append(SITE3[i][0])
        except IndexError:
            pass
        try:
            site.append(SITE4[i][0])
        except IndexError:
            pass

        try:
            site.append(SITE5[i][0])
        except IndexError:
            pass
        comentarios = COMENTARIOS[i][0]
        lider = LIDER[i][0]
        actividad = ACTIVIDAD[i][0]
        elementos = ELEMENTOS[i][0]

        #Cargados
        escenario = "Actividades de Cruzadas (Transmisión, migraciones, conexiones, pruebas, levantamiento de canalizac.)" #Constante
        tel_solicitante = solicitantes[solicitante][0]
        mail = solicitantes[solicitante][1]
        area = areas[approver]
        file_path_cuadrilla = r'C:\Users\Usuario\Desktop\K2K\Automatización\{}.csv'.format(lider)

        # Cambia a la otra URL
        driver.get('https://10.67.106.100/site_access_v2/trabajos/form?inframe=1&acceso=18229') 

        find_and_send_keys("solicitante",solicitante)
        find_and_send_keys("telefonoContacto",tel_solicitante)
        find_and_send_keys("correoContacto",mail)
        find_element("areaSolicitante",area)
        find_element("PreAprobador",approver)
        find_and_send_keys("personaComcel",approver)
        find_and_send_keys("telefonoPersonaComcel",tel_approver)
        find_and_click_by_value("afectaServicio", "NO")
        find_and_send_keys("comentariosTrabajo",comentarios)
        find_and_send_keys('fileMinutoGrama',file_path_minutograma)
        find_and_send_keys('fileOrdenCompra',file_path_poliza)
        find_and_send_keys('fileSoporteOrdenCompra',file_path_poliza)
        find_and_click("btnGuardar")
        for sitio in site:
            tipo = sitios[sitio]
            find_element("tipositios",tipo)
            time.sleep(1)
            find_element("combositios",sitio)
            find_element("comboelementos",elementos)
            find_and_click("btnAgregar")
        find_and_click("btnGuardar")
        find_and_click("btnConfirmar")
        find_and_send_keys('filePersonal',file_path_cuadrilla)
        find_and_click_upload()
        time.sleep(3.5)
        find_and_click("btnGuardar")
        find_and_click("btnConfirmar")
        find_element("tipoActividad",actividad)
        find_element("escenarioRiesgo",escenario)
        find_and_click("reqsegsal_1")
        find_and_click("reqsegsal_2")
        find_and_click("reqsegsal_5")
        find_and_send_keys('archivo_1',file_path_poliza)
        find_and_send_keys('archivo_2',file_path_poliza)
        find_and_send_keys('archivo_5',file_path_minutograma)
        start = datetime.now()
        start = start + timedelta(minutes=5)
        end = start + timedelta(days=14)
        end = end.replace(hour=18, minute=0, second=0, microsecond=0)
        format_date_time = "%d/%m/%y %H:%M"
        find_and_send_date("fechaPropuestaStr", str(start.strftime(format_date_time)))
        find_and_send_date("fechaFinTrabajoStr", str(end.strftime(format_date_time)))
        find_and_click("btnGuardar")
        find_and_click("btnConfirmar")
        time.sleep(1)
        id_site = wait.until(EC.presence_of_element_located((By.ID, "modal-resp")))
        id_site = id_site.text
        match = re.search(r'trabajo <b>(\d+)</b>', id_site)
        if match:
            id_site = match.group(1)
        else:
            print("No se encontró el ID del trabajo")
        values_col_K = [[id_site]]
        result_col_K = append_values_to_spreadsheet(
            sheet, SPREADSHEET_ID, f'K{i+2}', values_col_K, sheet_name='App'
        )

