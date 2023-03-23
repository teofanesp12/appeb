# !/usr/bin/env python3
# -*- coding: utf-8 -*-
from tools.configure import Configure
from tools.system import libreoffice_write, libreoffice_calc, editortxt, explore, convert_to_pdf
from os import path
from reengajamentos.app import Requerimento

# definimos a raiz do diretorio
rootDirURL = path.dirname(__file__)

from tkinter import *
from tkinter import ttk
import os
from threading import Thread

class Application:
    def __init__(self, master=None):
        self.widget1 = Frame(master)
        self.widget1.pack()
        self.msg = Label(self.widget1, text="Reengajamento - Selecione alguma ação abaixo:")
        self.msg["font"] = ("Calibri", "12", "italic", "bold")
        self.msg.pack ()

        self.abrir_csv = Button(self.widget1)
        self.abrir_csv["text"] = "EDITAR PLANILHA"
        self.abrir_csv["font"] = ("Calibri", "12")
        # self.abrir_csv["width"] = 5
        self.abrir_csv["command"] = self.open_csv
        self.abrir_csv.pack (side=RIGHT)

        self.msg3 = Label(self.widget1, text=" | ")
        self.msg3.pack (side=RIGHT)

        self.abrir_odtDESPACHO = Button(self.widget1)
        self.abrir_odtDESPACHO["text"] = "Modelo DESPACHO"
        self.abrir_odtDESPACHO["font"] = ("Calibri", "12")
        self.abrir_odtDESPACHO["command"] = self.open_odtDESPACHO
        self.abrir_odtDESPACHO.pack (side=RIGHT)

        self.abrir_odtINFO = Button(self.widget1)
        self.abrir_odtINFO["text"] = "Modelo INFO"
        self.abrir_odtINFO["font"] = ("Calibri", "12")
        self.abrir_odtINFO["command"] = self.open_odtINFO
        self.abrir_odtINFO.pack (side=RIGHT)
        
        self.abrir_odtREQUERIMENTO = Button(self.widget1)
        self.abrir_odtREQUERIMENTO["text"] = "Modelo REQUERIMENTO"
        self.abrir_odtREQUERIMENTO["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.abrir_odtREQUERIMENTO["command"] = self.open_odtREQUERIMENTO
        self.abrir_odtREQUERIMENTO.pack (side=RIGHT)

        self.msg2 = Label(self.widget1, text=" | ")
        self.msg2.pack (side=RIGHT)
        
        self.runbttn = Button(self.widget1)
        self.runbttn["text"] = "GERAR ARQUIVO"
        self.runbttn["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.runbttn["command"] = self.run_def
        self.runbttn.pack (side=RIGHT)

        self.prontos = Button(self.widget1)
        self.prontos["text"] = "ABRIR PRONTOS"
        self.prontos["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.prontos["command"] = self.open_prontos
        self.prontos.pack (side=RIGHT)

        self.prontos = Button(self.widget1)
        self.prontos["text"] = "GERAR PDFs"
        self.prontos["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.prontos["command"] = self.pdfs_prontos
        self.prontos.pack (side=RIGHT)

        self.widget2 = Frame(master)
        self.widget2.pack()
        self.pb = ttk.Progressbar(self.widget2, orient="horizontal", mode="determinate", length=820)# 640
        self.pb.pack()
        self.info= Text(self.widget2, height=10,width=115)#80
        self.info.config(state= DISABLED)
        self.info.pack()
    def pdfs_prontos(self):
        Thread(target=self.pdfs_prontos_thread).start()
    def pdfs_prontos_thread(self):
        prontosURL = os.path.join(rootDirURL, "prontos")
        try:
            os.mkdir(prontosURL)
        except OSError as error:
            pass
        convert_to_pdf(prontosURL)
        explore(prontosURL)
    def open_csv(self):
        listaURL = os.path.join(rootDirURL, "reengajamentos", "lista.csv")
        libreoffice_calc(listaURL)
    def open_prontos(self):
        prontosURL = os.path.join(rootDirURL, "prontos")
        try:
            os.mkdir(prontosURL)
        except OSError as error:
            self.info.config(state= NORMAL)
            self.info.insert(INSERT, "arquivo já criado...\n")
            self.info.config(state= DISABLED)
        Thread(target=explore, args=(prontosURL,)).start()
    def open_odtREQUERIMENTO(self):
        odtURL = os.path.join(rootDirURL, "reengajamentos", "modelo_requerimento.odt")
        libreoffice_write(odtURL)
    def open_odtINFO(self):
        odtURL = os.path.join(rootDirURL, "reengajamentos", "modelo_info.odt")
        libreoffice_write(odtURL)
    def open_odtDESPACHO(self):
        odtURL = os.path.join(rootDirURL, "reengajamentos", "modelo_despacho.odt")
        libreoffice_write(odtURL)

    def run_def(self):
        Thread(target=self.run).start()
    def run(self):
        configure = Configure(rootDirURL)
        # definimos o modelo Alteração
        configure.setzipSourceFile("reengajamentos","modelo_requerimento.odt")
        # definimos o arquivo lista
        configure.setcsvSourceFile("reengajamentos","lista.csv")
        # definimos o arquivo de configuração
        # configure.setiniSourceFile("reengajamentos","configure.ini")
        #
        # iniciamos...
        #
        self.info.config(state= NORMAL)
        len_data = len(configure.getData())*3
        self.pb['value'] = 0
        i = 0
        for linha in configure.getData():
            self.info.insert(INSERT, linha[0])
            arquivo = Requerimento()
            arquivo.nome = "Requerimento "+linha[0]
            arquivo.setConfigure(configure)
            arquivo.set_data(linha, headers=configure.getData().headers)
            arquivo.extrair()
            arquivo.replace()
            arquivo.write()
            arquivo.comprimir()
            self.pb['value'] = i/len_data*100
            i+=1
            self.info.insert(INSERT, " - arquivo REQUERIMENTO pronto com sucesso\n\n")

        configure.setzipSourceFile("reengajamentos","modelo_info.odt")
        for linha in configure.getData():
            self.info.insert(INSERT, linha[0])
            arquivo = Requerimento()
            arquivo.nome = "Info "+linha[0]
            arquivo.setConfigure(configure)
            arquivo.set_data(linha, headers=configure.getData().headers)
            arquivo.extrair()
            arquivo.replace()
            arquivo.write()
            arquivo.comprimir()
            self.pb['value'] = i/len_data*100
            i+=1
            self.info.insert(INSERT, " - arquivo INFO pronto com sucesso\n\n")

        configure.setzipSourceFile("reengajamentos","modelo_despacho.odt")
        for linha in configure.getData():
            self.info.insert(INSERT, linha[0])
            arquivo = Requerimento()
            arquivo.nome = "Despacho "+linha[0]
            arquivo.setConfigure(configure)
            arquivo.set_data(linha, headers=configure.getData().headers)
            arquivo.extrair()
            arquivo.replace()
            arquivo.write()
            arquivo.comprimir()
            self.pb['value'] = i/len_data*100
            i+=1
            self.info.insert(INSERT, " - arquivo DESPACHO pronto com sucesso\n\n")
        
        self.info.config(state= DISABLED)
        self.pb['value'] = 100

root = Tk()
Application(root)
root.title('Reengajamento de CBs e SDs')
root.iconphoto(False, PhotoImage(file = os.path.join(rootDirURL, "static", "militar.png")))
root.mainloop()
