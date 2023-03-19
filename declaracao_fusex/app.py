# -*- coding: utf-8 -*-
import sys
sys.path.append("..")

import os
import zipfile
import fileinput
import tablib

from tools.odt import ArquivoODT

if __name__ == "__main__":
    from tools.configure import Configure
    from os import path

    # definimos a raiz do diretorio
    rootDirURL = path.dirname(__file__)
    configure = Configure(rootDirURL)
    # definimos o modelo Alteração
    configure.setzipSourceFile("modelo.odt")
    # definimos o arquivo lista
    configure.setcsvSourceFile("lista.csv")

    for linha in configure.getData():
        arquivo = ArquivoODT()
        arquivo.nome = "Declaração FuSEx "+linha[0]
        arquivo.setConfigure(configure)
        arquivo.set_data(linha, headers=configure.getData().headers)
        arquivo.extrair()
        arquivo.replace()
        arquivo.write()
        arquivo.comprimir()

