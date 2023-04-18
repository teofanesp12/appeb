# -*- coding: utf-8 -*-
import sys
sys.path.append("..")

import os
import zipfile
import fileinput
import tablib

from tools.system import convert_to_pdf
from tools.odt import ArquivoODT

class CartaoCabelo(ArquivoODT):
    def __init__(self):
        super().__init__()
    def extrair(self):
        #
        # Unzip ODT
        #
        tmpDir="~odt_contents"+self.nome
        self._tmpDirURL=path.join(self._configure.rootDirURL, "tmp", tmpDir)
        info (self._configure.tk, "")
        info (self._configure.tk," -- Extraindo ---------------------")
        info (self._configure.tk,"%s -> %s" % (self._configure.zipSourceFileURL, self._tmpDirURL))

        self._zipdata = zipfile.ZipFile(self._configure.zipSourceFileURL)
        self._zipdata.extractall(self._tmpDirURL)

        # abrimos os arquivos principais...
        
        xmlFile1="content.xml"
        xmlFileURL1=path.join(self._tmpDirURL, 'Object 2', xmlFile1)
        self.content = Content(xmlFileURL1)
        
        xmlFile2="styles.xml"
        xmlFileURL2=path.join(self._tmpDirURL, 'Object 2', xmlFile2)
        self.styles = Styles(xmlFileURL2)
    def comprimir(self):
        zipOutFile=self.nome+".odt"
        zipOutFileURL=path.join(self._configure.prontos, zipOutFile)
        try:
            os.mkdir(self._configure.prontos)
        except OSError as error:
            info(self._configure.tk,"Se arquivo já foi criado ignore, caso contrario confira o acesso da pasta 'prontos'")
        info (self._configure.tk," -- Comprimindo --------------------")
        info (self._configure.tk,"%s -> %s" % (self._tmpDirURL , zipOutFileURL))
        info (self._configure.tk,"")

        with zipfile.ZipFile(zipOutFileURL, 'w') as outzip:
            zipinfos = self._zipdata.infolist()
            for zipinfo in zipinfos:
                fileName=zipinfo.filename # The name and path as stored in the archive
                fileURL=path.join(self._tmpDirURL, fileName)# The actual name and path
                outzip.write(fileURL,fileName)

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
        arquivo = CartaoCabelo()
        arquivo.nome = linha[0]
        arquivo.setConfigure(configure)
        arquivo.set_data(linha, headers=configure.getData().headers)
        arquivo.extrair()
        arquivo.replace()
        arquivo.write()
        arquivo.comprimir()
    prontosURL = os.path.join(rootDirURL, "prontos")
    convert_to_pdf(prontosURL)
