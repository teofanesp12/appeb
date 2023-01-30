# -*- coding: utf-8 -*-
import sys
sys.path.append("..")

import os
import zipfile
import fileinput
import tablib

from tools.odt import ArquivoODT

class Requerimento(ArquivoODT):
    def __init__(self):
        super().__init__()

    def replace2(self):
        #
        # Find and replace tokens
        #
        xmlFileURL = self.content._arquivo
        print (" -- Replacing -------------")
        print (xmlFileURL)

        # xmlFileObj = open(xmlFileURL, "r+", encoding="utf-8")
        for line in fileinput.FileInput(xmlFileURL, inplace=1):
            print (line.replace("{nome}",self._data['nome']).replace("{nup}",self._data['nup']).replace("{identidade}",self._data['identidade']).replace("{requer}",self._data['requer']))

if __name__ == "__main__":
    from tools.configure import Configure
    from os import path

    # definimos a raiz do diretorio
    rootDirURL = path.dirname(__file__)
    configure = Configure(rootDirURL)
    # definimos o modelo Alteração
    configure.setzipSourceFile("modelo_requerimento.odt")
    # definimos o arquivo lista
    configure.setcsvSourceFile("lista.csv")

    for linha in configure.getData():
        arquivo = Requerimento()
        arquivo.nome = "Requerimento "+linha[0]
        arquivo.setConfigure(configure)
        arquivo.set_data(linha, headers=configure.getData().headers)
        arquivo.extrair()
        arquivo.replace()
        arquivo.write()
        arquivo.comprimir()

    configure.setzipSourceFile("modelo_info.odt")
    for linha in configure.getData():
        arquivo = ArquivoODT()
        arquivo.nome = "Info "+linha[0]
        arquivo.setConfigure(configure)
        arquivo.set_data(linha, headers=configure.getData().headers)
        arquivo.extrair()
        arquivo.replace()
        arquivo.write()
        arquivo.comprimir()

    configure.setzipSourceFile("modelo_despacho.odt")
    for linha in configure.getData():
        arquivo = ArquivoODT()
        arquivo.nome = "Despacho "+linha[0]
        arquivo.setConfigure(configure)
        arquivo.set_data(linha, headers=configure.getData().headers)
        arquivo.extrair()
        arquivo.replace()
        arquivo.write()
        arquivo.comprimir()
