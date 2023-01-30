# -*- coding: utf-8 -*-
import os

def libreoffice_write(file_url):
    if os.name == 'nt':
        os.system('start "C:\\Program Files\\LibreOffice\\program\\swriter.exe" "%s"'%file_url)
    elif os.name == 'posix':
        os.system('libreoffice "%s"'%file_url)
def libreoffice_calc(file_url):
    if os.name == 'nt':
        os.system('start "C:\\Program Files\\LibreOffice\\program\\scalc.exe" "%s"'%file_url)
    elif os.name == 'posix':
        os.system('libreoffice "%s"'%file_url)

def editortxt(file_url):
    if os.name == 'nt':
        os.system('start "C:\\Windows\\notepad.exe" "%s"'%file_url)
    elif os.name == 'posix':
        os.system('gedit "%s"'%file_url)

def explore(file_url):
    if os.name == 'nt':
        os.system('start "C:\\Windows\\explorer.exe" "%s"'%file_url)
    elif os.name == 'posix':
        os.system('nautilus "%s"'%file_url)
