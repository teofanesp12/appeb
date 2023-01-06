from tools.configure import Configure
from os import path
from reengajamentos.app import Requerimento

# definimos a raiz do diretorio
rootDirURL = path.dirname(__file__)
configure = Configure(rootDirURL)
# definimos o modelo Alteração
configure.setzipSourceFile("reengajamentos\\requerimento_modelo.odt")
# definimos o arquivo lista
configure.setcsvSourceFile("reengajamentos\\lista.csv")
# definimos o arquivo de configuração
# configure.setiniSourceFile("reengajamentos\\configure.ini")

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
        
        self.runbttn = Button(self.widget1)
        self.runbttn["text"] = "GERAR ARQUIVO"
        self.runbttn["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.runbttn["command"] = self.run
        self.runbttn.pack (side=RIGHT)

    def open_csv(self):
        os.system('start "C:\\Program Files\\LibreOffice\\program\\scalc.exe" "%s"'%rootDirURL+"\\reengajamentos\\lista.csv")
    def open_odt(self):
        os.system('start "C:\\Program Files\\LibreOffice\\program\\swriter.exe" "%s"'%rootDirURL+"\\reengajamentos\\requerimento_modelo.odt")

    def run(self):
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
            arquivo.comprimir()

root = Tk()
Application(root)
root.title('Reengajamento de CBs e SDs')
root.mainloop()
