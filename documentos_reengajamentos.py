#! /usr/bin/python
from tools.configure import Configure
from tools.system import libreoffice_write, libreoffice_calc, editortxt, explore
from os import path
from reengajamentos.app import Requerimento

# definimos a raiz do diretorio
rootDirURL = path.dirname(__file__)

from tkinter import *
import os

class Application:
    def __init__(self, master=None):
        self.widget1 = Frame(master)
        self.widget1.pack()
        self.msg = Label(self.widget1, text="Reengajamento - Selecione alguma ação abaixo:")
        self.msg["font"] = ("Calibri", "12", "italic", "bold")
        self.msg.pack ()

        self.abrir_csv = Button(self.widget1)
        self.abrir_csv["text"] = "EDITAR Arquivo CSV"
        self.abrir_csv["font"] = ("Calibri", "12")
        # self.abrir_csv["width"] = 5
        self.abrir_csv["command"] = self.open_csv
        self.abrir_csv.pack (side=RIGHT)
        
        self.abrir_odt = Button(self.widget1)
        self.abrir_odt["text"] = "EDITAR Arquivo de Modelo"
        self.abrir_odt["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.abrir_odt["command"] = self.open_odt
        self.abrir_odt.pack (side=RIGHT)

        self.msg2 = Label(self.widget1, text=" | ")
        self.msg2.pack (side=RIGHT)
        
        self.runbttn = Button(self.widget1)
        self.runbttn["text"] = "GERAR ARQUIVO"
        self.runbttn["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.runbttn["command"] = self.run
        self.runbttn.pack (side=RIGHT)

        self.prontos = Button(self.widget1)
        self.prontos["text"] = "ABRIR PRONTOS"
        self.prontos["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.prontos["command"] = self.open_prontos
        self.prontos.pack (side=RIGHT)

    def open_csv(self):
        listaURL = os.path.join(rootDirURL, "reengajamentos", "lista.csv")
        libreoffice_calc(listaURL)
    def open_prontos(self):
        prontosURL = os.path.join(rootDirURL, "prontos")
        try:
            os.mkdir(prontosURL)
        except OSError as error:
            print("arquivo já criado...")
        explore(prontosURL)
    def open_odt(self):
        odtURL = os.path.join(rootDirURL, "reengajamentos", "requerimento_modelo.odt")
        libreoffice_write(odtURL)

    def run(self):
        configure = Configure(rootDirURL)
        # definimos o modelo Alteração
        configure.setzipSourceFile("reengajamentos","requerimento_modelo.odt")
        # definimos o arquivo lista
        configure.setcsvSourceFile("reengajamentos","lista.csv")
        # definimos o arquivo de configuração
        # configure.setiniSourceFile("reengajamentos","configure.ini")
        #
        # iniciamos...
        #
        for linha in configure.getData():
            arquivo = Requerimento()
            arquivo.nome = linha[0]
            arquivo.setConfigure(configure)
            arquivo.set_data(linha, headers=configure.getData().headers)
            arquivo.extrair()
            arquivo.replace()
            arquivo.write()
            arquivo.comprimir()

root = Tk()
Application(root)
root.title('Reengajamento de CBs e SDs')
root.mainloop()
