# !/usr/bin/env python3
from tools.configure import Configure
from tools.system import libreoffice_write, libreoffice_calc, editortxt, explore

from os import path
from assentamentos.app import ArquivoODT

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
        self.msg = Label(self.widget1, text="Alterações e Assentamentos - Selecione alguma ação abaixo:")
        self.msg["font"] = ("Calibri", "12", "italic", "bold")
        self.msg.pack ()

        self.abrir_csv = Button(self.widget1)
        self.abrir_csv["text"] = "EDITAR Arquivo Lista"
        self.abrir_csv["font"] = ("Calibri", "12")
        # self.abrir_csv["width"] = 5
        self.abrir_csv["command"] = self.open_csv
        self.abrir_csv.pack (side=RIGHT)
        
        self.abrir_ini = Button(self.widget1)
        self.abrir_ini["text"] = "EDITAR Arquivo de Configuração"
        self.abrir_ini["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.abrir_ini["command"] = self.open_ini
        self.abrir_ini.pack (side=RIGHT)

        self.msg2 = Label(self.widget1, text=" | ")
        self.msg2.pack (side=RIGHT)
        
        self.runbttn = Button(self.widget1)
        self.runbttn["text"] = "GERAR ARQUIVOS"
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

        self.widget2 = Frame(master)
        self.widget2.pack()
        self.pb = ttk.Progressbar(self.widget2, orient="horizontal", mode="determinate", length=580)
        self.pb.pack()
        self.info= Text(self.widget2, height= 10,width= 80)
        self.info.config(state= DISABLED)
        self.info.pack()

    def open_csv(self):
        listaURL = os.path.join(rootDirURL, "assentamentos", "lista.csv")
        libreoffice_calc(listaURL)
    def open_prontos(self):
        prontosURL = os.path.join(rootDirURL, "prontos")
        try:
            os.mkdir(prontosURL)
        except OSError as error:
            self.info.config(state= NORMAL)
            self.info.insert(INSERT, "arquivo já criado...\n")
            self.info.config(state= DISABLED)
            # print("arquivo já criado...")
        explore(prontosURL)
    def open_ini(self):
        configureURL = os.path.join(rootDirURL, "assentamentos", "configure.ini")
        editortxt(configureURL)

    def run_def(self):
        Thread(target=self.run).start()
    def run(self):
        configure = Configure(rootDirURL)
        # definimos o modelo Alteração
        configure.setzipSourceFile("assentamentos","modelo_alteracao.odt")
        # definimos o arquivo lista
        configure.setcsvSourceFile("assentamentos","lista.csv")
        # definimos o arquivo de configuração
        configure.setiniSourceFile("assentamentos","configure.ini")
        #
        # iniciamos...
        #
        self.info.config(state= NORMAL)
        len_data = len(configure.getData())
        self.pb['value'] = 0
        i = 0
        for linha in configure.getData():
            self.info.insert(INSERT, linha[0])
            arquivo = ArquivoODT()
            arquivo.nome = linha[0]
            arquivo.setConfigure(configure)
            arquivo.set_data(linha, headers=configure.getData().headers)
            # self.info.insert(INSERT, ": Iniciou com sucesso")
            arquivo.extrair()
            arquivo.run()
            arquivo.comprimir()
            self.info.insert(INSERT, " - arquivo ODT pronto com sucesso\n\n")
            self.pb['value'] = i/len_data*100
            i+=1
        self.info.config(state= DISABLED)
        self.pb['value'] = 100

root = Tk()
Application(root)
root.title('Alterações e Assentamentos')
root.mainloop()
