# Misceláneas
import re
import base64
import glob
from sys import exit
import pandas as pd
import os
import shutil
import json
from datetime import datetime
import matplotlib.pyplot as plt
import subprocess
from openai import OpenAI
import fitz
from openpyxl import load_workbook
import paramiko
import requests
from msal import ConfidentialClientApplication
scope = ["https://graph.microsoft.com/.default"]

# Correo
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from os.path import basename

# Sharepoint
from shareplum import Site
from shareplum import Office365
from shareplum.site import Version
from datetime import datetime
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# Web Scraping y paralelización
from selenium.webdriver.common.keys import Keys
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support import expected_conditions as EC
import time
import threading

def espera(driver,tiempo,com,excep=""):
    """Make the driver wait for specified seconds (at most) while executing an action

    Parameters
    ----------
    driver : WebDriver
        The driver what is being used
    tiempo : int
        The time (in seconds) multiplied by 5 is the time will the driver wait at most (checking each 5 seconds)
    com : str
        The action the driver will try
    excep : str
        The action that will be an exception (For "Athena_Medical_Records" only)
    Returns
    -------
    None
    """
    a = True
    p = 0
    while a==True and p<=tiempo:
        try:
            exec(com)
            a = False
            #print("al fin...")
        except:
            if excep!="":
                try:
                    exec(excep)
                    if len(b)==1:
                        a=False
                        return False
                except:
                    pass
            time.sleep(5)
            #print("esperando")
            p += 1
            pass
            
def dividir_lista(lista, n):
    """Make the list be divided in lenght of the list by n (approximately) for getting elements for each window in driver
    if the list have 12 elements -> [1, 2, 3, ..., 12]
    Example 1: and n = 3, it will create 4 parts -> 4 parts with 3 elements
    Example 2: and n = 4, it will create 3 parts -> 3 parts with 4 elements
    Example 3: and n = 5, it will create 3 parts -> 2 parts with 5 elements and 1 with 2 elements

    Parameters
    ----------
    lista : list
        The input list with the total elements
    n : int
        The number that will divide the total elements

    Returns
    -------
    The parts calculated
        Note: one part = one window for driver
    """
    k, m = divmod(len(lista), n)
    return (lista[i * k + min(i, m):(i + 1) * k + min(i + 1, m)] for i in range(n))

def dividir_diccionario(dic, n):
    """Make the dic be divided in the number of keys (with its elements) by n (approximately) for getting elements for each window in driver
    if the dictionary have 12 keys -> {"key1": [1,2,...], "key2": [3,6,...], ..., "key12": [9,1,...]}
    Example 1: and n = 3, it will create 4 parts -> 4 parts with 3 keys and its respective elements
    Example 2: and n = 4, it will create 3 parts -> 3 parts with 4 keys and its respective elements
    Example 3: and n = 5, it will create 3 parts -> 2 parts with 5 keys and its respective elements and 1 part with 2 keys and its respective elements

    Parameters
    ----------
    dic : dictionary
        The input list with the total elements
    n : int
        The number that will divide the total elements

    Returns
    -------
    The parts calculated
        Note: one part = one window for driver
    """
    keys = list(dic.keys())
    division = len(keys) // n
    if division==0:
      division=1
    print(division)
    return [dict((k, dic[k]) for k in keys[i:i + division]) for i in range(0, len(keys), division)]

def send_email(address, password, subject, text_in, files_list=[], emails_list=[]):
    """Send an email (Gmail only) depending if there are files listed or not
    Parameters
    ----------
    address: str
        Gmail that will send the email
    password: str
        third party password that is created in https://myaccount.google.com/apppasswords (remember to have 2 step pass active and remember your API key)
    subject: str
        subject for email
    text_in: str
        text to write in email
    files_list: list
        Files that will be send
    emails_list: list
        List of emails that will recieve the email sent

    Returns
    -------
    None
    """
    
    sender_address = address
    sender_pass = password
    session = smtplib.SMTP('smtp.gmail.com',587)
    #session = smtplib.SMTP('smtp.office365.com',587)
    session.starttls()
    session.login(sender_address, sender_pass)
    files=files_list
    
    #Sending
    for correo in emails_list:
      message=MIMEMultipart()
      message['From'] = sender_address
      if len(files_list)>0:
        message['Subject'] = subject
        mail_content = text_in
        message.attach(MIMEText(mail_content,'plain'))
        for f in files or []:
            with open(f, "rb") as fil:
                ext = f.split('.')[-1:]
                attachedfile = MIMEApplication(fil.read(), _subtype = ext)
                attachedfile.add_header(
                    'content-disposition', 'attachment', filename=basename(f) )
            message.attach(attachedfile)
      else:
        message['Subject'] = subject
        mail_content = text_in
        message.attach(MIMEText(mail_content,'plain'))
      text = message.as_string()
      session.sendmail(sender_address, correo, text)
      print("Correo enviado")
    session.quit()

def download_sharepoint(ruta,file,site):
    """Download file from SharePoint
    Parameters
    ----------
    ruta: str
        Path where file is stored
    password: str
        File that will be downloaded
    site: shareplum.site object
        Site where Shareplum is focused to download

    Returns
    -------
    None
    """
    folder = site.Folder(ruta)
    local_file_path = file
    file_content = folder.get_file(file)
    with open(local_file_path, 'wb') as file:
        file.write(file_content)

def upload_sharepoint(url,nombres_archivos,user_sharepoint,contra_sharepoint,relative_url,go=0):
    """Upload file to SharePoint
    Parameters
    ----------
    url: str
        SharePoint origin URL and path
    nombres_archivos: list
        List of files to upload
    user_sharepoint: str
        SharePoint username
    contra_sharepoint: str
        Sharepoint password
    relative_url: str
        Aditional path to url
    go: int
        Check if ignore file that contains " - 20" in name

    Returns
    -------
    None
    """
    ctx_auth = AuthenticationContext(url)
    for file_path in nombres_archivos:
        if go!=0:
            if " - 20" in file_path:
                print("ignorando")
                continue       
        if ctx_auth.acquire_token_for_user(user_sharepoint, contra_sharepoint):
            ctx = ClientContext(url, ctx_auth)
            with open(file_path, 'rb') as content_file:
                file_content = content_file.read()
            dir, name = os.path.split(file_path)
            target_folder = ctx.web.get_folder_by_server_relative_url(relative_url)
            target_file = target_folder.upload_file(name, file_content).execute_query()
            print(f"Archivo {name} subido a {target_folder.serverRelativeUrl}")
        else:
            print("Error en la autenticación")

def opciones_driver(a=0):
    """Send an email (Gmail only) depending if there are files listed or not
    Parameters
    ----------
    None

    Returns
    -------
    a: WebDriver Headless
        0 headless, not 0 not headless
    """
    # Web driver
    options = Options()
    if a==0:
        options.add_argument('--headless') # Avoid colab to crash
    else:
        pass
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage') # Avoid unexpected errors
    options.add_argument('--window-size=1920,1080') # Specify window size
    options.add_argument("--incognito") # Necessary for parallelization
    ruta_descargas = os.getcwd()+"\Athena_reports"
    # Set download route
    try:
        os.mkdir(ruta_descargas)
    except:
        pass
    prefs = {
        "download.default_directory": ruta_descargas,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return options, ruta_descargas

def opciones_driver_many(k,a=0):
    """Send an email (Gmail only) depending if there are files listed or not
    Parameters
    ----------
    None

    Returns
    -------
    k: Webdriver window
        any int
    a: WebDriver Headless
        0 headless, not 0 not headless
    """
    # Web driver
    options = Options()
    if a==0:
        options.add_argument('--headless') # Avoid colab to crash
    else:
        pass
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage') # Avoid unexpected errors
    options.add_argument('--window-size=1920,1080') # Specify window size
    options.add_argument("--incognito") # Necessary for parallelization
    ruta_descargas = os.getcwd()+"\Athena_reports"
    # Set download route
    try:
        os.mkdir(ruta_descargas)
    except:
        pass
    prefs = {
        "download.default_directory": os.path.join(ruta_descargas, str(k)),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    return options, ruta_descargas

def athena_enter(k,ruta_descargas,options,user,contra):
    """Enter to Athena website
    Parameters
    ----------
    k: int
        number of window
    ruta_descargas: str
        Download path
    options: WebDriver options
        Webdriver options
    user: str
        Athena user
    contra: str
        Athena password

    Returns
    -------
    driver: WebDriver
        Webdriver with Athena website in
    """
    time.sleep(k*10)
    print("Entering with window #", k)
    driver = webdriver.Edge(options=options)
    #print("va")
    os.chdir(ruta_descargas)
    #print("va2")
    a = True
    red=0
    while a==True and red<=2:
      try:
          driver.get('https://athenanet.athenahealth.com/')
          #driver.get('https://identity.athenahealth.com/oauth2/auset0ja9xZ2Hniep296/v1/authorize?client%5Fid=0oaet0rfjNzyKCiQQ296&idp=&login%5Fhint=ANETUSERNAME&nonce=031d476cb678c4a0d9dd13b1634d5d492077c5704c0ad3832c9ae95214808934&prompt=login&redirect%5Furi=https%3A%2F%2Fathenanet%2Eathenahealth%2Ecom%2F1%2F1%2Flogin%2Foidc%2Eesp&response%5Fmode=form%5Fpost&response%5Ftype=code&scope=openid%20profile%20offline%5Faccess%20&sessionToken=&state=eyJGTEFHUyI6eyJDT0RFUEFTU1RIUk9VR0giOm51bGwsIkFORVRNRkFTSElNV0lER0VUIjoiIiwiTk9GUkFNRVNFVCI6bnVsbCwiREVQQVJUTUVOVElEIjpudWxsLCJERUVQTElOSyI6bnVsbH0sIkxPR0lOTUVUQURBVEEiOnsiVVNFUkFVVEhOVFlQRSI6Ik5PTlNTTyIsIkJBTk5FUlRZUEUiOiJsaXZlX2xvZ2luIn0sIkNTUkYiOiJlODEzZWQxMWY3M2ViMzllNTQwMWE0ZjdjM2ZmZGMwMSIsIlRBUkdFVFVSTCI6bnVsbH0')
          #print("yea", k)
          time.sleep(10)
          #driver.save_screenshot("aqui.png")
          #plt.imshow(plt.imread("aqui.png"))
          driver.find_element(By.XPATH, '//input[@id="athena-username"]').send_keys(user)
          driver.find_element(By.XPATH, '//input[@id="athena-password"]').send_keys(contra)
          driver.find_element(By.XPATH, '//button[@type="button"]').click()
          time.sleep(5)
          #driver.save_screenshot(f"aqui_{k}.png")
          #plt.imshow(plt.imread("aqui.png"))
          driver.find_element(By.XPATH, '//input[@id="loginbutton"]').click()
          print("Go #", k)
          a = False
      except:
          print("Oh no", k)
          driver.quit()
          time.sleep(5)
          red+=1
          driver = webdriver.Edge(options=options)
    return driver

def optimum_enter(k,ruta_descargas,options,user,contra):
    """Enter to Optimum website
    Parameters
    ----------
    k: int
        number of window
    ruta_descargas: str
        Download path
    options: WebDriver options
        Webdriver options
    user: str
        Optimum user
    contra: str
        Optimum password

    Returns
    -------
    driver: WebDriver
        Webdriver with Athena website in
    """
    time.sleep(k*10)
    print("Entering with window #", k)
    driver = webdriver.Edge(options=options)
    os.chdir(ruta_descargas)
    driver.get('https://optprovider.prod.healthaxis.net/login/')
    time.sleep(10)
    driver.find_element(By.XPATH, '//input[@name="userName"]').send_keys(user)
    driver.find_element(By.XPATH, '//input[@name="password"]').send_keys(contra)
    driver.find_element(By.XPATH, '//button[@id="submitLoginForm"]').click()
    time.sleep(5)    
    driver.find_element(By.XPATH, '//button[@ng-class="okBtnClass"]').click()
    time.sleep(1)
    driver.refresh()
    time.sleep(1)
    a=driver.find_elements(By.XPATH,"//a[@data-target='#authorization']")
    a[1].click()
    a=driver.find_elements(By.XPATH,"//a[@ui-sref='claimSearch()']")
    a[0].click()
    time.sleep(3)
    return driver

class SharePointClient:
    def __init__(self, client_id, authority, client_secret, scope, sharepoint_site_url):
        self.client_id = client_id
        self.authority = authority
        self.client_secret = client_secret
        self.scope = scope
        self.sharepoint_site_url = sharepoint_site_url
        self.token_acceso = self.obtener_token_acceso()
        self.site_id = self.obtener_site_id()
        self.drive_id = self.obtener_drive_id()

    # Método para obtener el token de acceso
    def obtener_token_acceso(self):
        app = ConfidentialClientApplication(self.client_id, authority=self.authority, client_credential=self.client_secret)
        result = app.acquire_token_for_client(scopes=self.scope)
        if "access_token" in result:
            return result["access_token"]
        else:
            raise Exception("Error al obtener el token de acceso: ", result)

    # Método para obtener el ID del sitio de SharePoint
    def obtener_site_id(self):
        url = f"https://graph.microsoft.com/v1.0/sites/{self.sharepoint_site_url}"
        headers = {
            'Authorization': f'Bearer {self.token_acceso}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            site_info = response.json()
            return site_info['id']
        else:
            raise Exception(f"Error al obtener el ID del sitio: {response.status_code} - {response.text}")

    # Método para obtener el ID de la biblioteca de documentos (drive_id)
    def obtener_drive_id(self):
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives"
        headers = {
            'Authorization': f'Bearer {self.token_acceso}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            drives_info = response.json()
            for drive in drives_info['value']:
                if drive['name'] == "Documents":
                    return drive['id']
            raise Exception("No se encontró la biblioteca de documentos 'Documents'")
        else:
            raise Exception(f"Error al obtener la biblioteca de documentos: {response.status_code} - {response.text}")

    # Método para obtener la lista de archivos en SharePoint
    def get_sharepoint_filenames(self, folder_path):
        url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.drive_id}/root:/{folder_path}:/children"
        headers = {
            'Authorization': f'Bearer {self.token_acceso}',
            'Content-Type': 'application/json'
        }
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            files_info = response.json()
            return [item['name'] for item in files_info['value']]
        else:
            raise Exception(f"Error al obtener archivos de SharePoint: {response.status_code} - {response.text}")

    # Método para subir archivos a SharePoint
    def upload_sharepoint_file(self, folder_path, nombre_archivo):
        #ruta_local = os.getcwd()
        with open(nombre_archivo, 'rb') as file_content:
            file_data = file_content.read()
        upload_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.drive_id}/root:/{folder_path}/{nombre_archivo}:/content"
        headers = {
            'Authorization': f'Bearer {self.token_acceso}',
            'Content-Type': 'application/octet-stream'
        }
        response = requests.put(upload_url, headers=headers, data=file_data)
        if response.status_code in [200, 201]:
            print(f"Archivo '{nombre_archivo}' subido correctamente a SharePoint.")
            #os.remove(ruta_local)  # Eliminar el archivo local después de subirlo
        else:
            raise Exception(f"Error al subir el archivo '{nombre_archivo}': {response.status_code} - {response.text}")

    # Método para descargar archivos de SharePoint
    def download_sharepoint_file(self, folder_path, nombre_archivo):
        ruta_local = os.path.join(os.getcwd(), nombre_archivo)
        download_url = f"https://graph.microsoft.com/v1.0/sites/{self.site_id}/drives/{self.drive_id}/root:/{folder_path}/{nombre_archivo}:/content"
        headers = {
            'Authorization': f'Bearer {self.token_acceso}'
        }
        response = requests.get(download_url, headers=headers, stream=True)
        if response.status_code == 200:
            with open(ruta_local, 'wb') as file:
                for chunk in response.iter_content(chunk_size=1024):
                    if chunk:
                        file.write(chunk)
            print(f"Archivo '{nombre_archivo}' descargado exitosamente a '{ruta_local}'.")
        else:
            raise Exception(f"Error al descargar el archivo '{nombre_archivo}': {response.status_code} - {response.text}")

