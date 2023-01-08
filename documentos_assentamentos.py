from tools.configure import Configure
from os import path
from assentamentos.app import ArquivoODT

# definimos a raiz do diretorio
rootDirURL = path.dirname(__file__)
configure = Configure(rootDirURL)
# definimos o modelo Alteração
configure.setzipSourceFile("assentamentos\\modelo_alteracao.odt")
# definimos o arquivo lista
configure.setcsvSourceFile("assentamentos\\lista.csv")
# definimos o arquivo de configuração
configure.setiniSourceFile("assentamentos\\configure.ini")

from tkinter import *
import os

class Application:
    def __init__(self, master=None):
        self.widget1 = Frame(master)
        self.widget1.pack()
        self.msg = Label(self.widget1, text="Alterações e Assentamentos - Selecione alguma ação abaixo:")
        self.msg["font"] = ("Calibri", "12", "italic", "bold")
        self.msg.pack ()

        self.abrir_csv = Button(self.widget1)
        self.abrir_csv["text"] = "EDITAR Arquivo CSV"
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
        self.runbttn["command"] = self.run
        self.runbttn.pack (side=RIGHT)

        self.prontos = Button(self.widget1)
        self.prontos["text"] = "ABRIR PRONTOS"
        self.prontos["font"] = ("Calibri", "12")
        # self.abrir_ini["width"] = 5
        self.prontos["command"] = self.open_prontos
        self.prontos.pack (side=RIGHT)

        self.widget2 = Frame(master)
        self.widget2.pack()

    def open_csv(self):
        os.system('start "C:\\Program Files\\LibreOffice\\program\\scalc.exe" "%s"'%rootDirURL+"\\assentamentos\\lista.csv")
    def open_prontos(self):
        try:
            os.mkdir(rootDirURL+"\\prontos")
        except OSError as error:
            print(error)
        os.system('start "C:\\Windows\\explorer.exe" "%s"'%rootDirURL+"\\prontos")
    def open_ini(self):
        os.system('start "C:\\Windows\\notepad.exe" "%s"'%rootDirURL+"\\assentamentos\\configure.ini")

    def run(self):
        #
        # iniciamos...
        #
        for linha in configure.getData():
            arquivo = ArquivoODT()
            arquivo.nome = linha[0]
            arquivo.setConfigure(configure)
            arquivo.set_data(linha, headers=configure.getData().headers)
            arquivo.extrair()
            arquivo.run()
            arquivo.comprimir()

root = Tk()
Application(root)
root.title('Alterações e Assentamentos')
root.mainloop()
