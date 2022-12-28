from lxml import etree
import zipfile

import os
from os import path

namespace = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
NSMAP = {
    "text":"urn:oasis:names:tc:opendocument:xmlns:text:1.0",
    "table":"urn:oasis:names:tc:opendocument:xmlns:table:1.0",
    "office":"urn:oasis:names:tc:opendocument:xmlns:office:1.0",
    "style":"urn:oasis:names:tc:opendocument:xmlns:style:1.0",
}

def _insert_namespace(value, nsmap='text'):
    return "{%s}%s"%(NSMAP[nsmap],value)
def new_element(value, text = None, attrs=[], nsmap='text'):
    el = etree.Element(_insert_namespace(value, nsmap=nsmap), nsmap = NSMAP)
    for attr in attrs:
        if len(attr)==2:
            el.set(_insert_namespace(attr[0], nsmap=nsmap), attr[1])
        elif len(attr)==3:
            el.set(_insert_namespace(attr[0], nsmap=attr[2]), attr[1])
    if text:
        el.text = text
    return el


class Content:
    def __init__(self, arquivo=""):
        self.scripts = ""
        self.font = ""
        self.styles = ""
        self.body = ""
        self.content = None
        if arquivo:
            self.read(arquivo)

    def read(self, arquivo):
        self._arquivo = arquivo
        self.content = etree.parse(arquivo)
        for child in self.content.getroot():
            if child.tag == "{%s}%s"%(namespace,"scripts"):
                self.scripts = child
            elif child.tag == "{%s}%s"%(namespace,"font-face-decls"):
                self.font = child
            elif child.tag == "{%s}%s"%(namespace,"automatic-styles"):
                self.styles = child
            elif child.tag == "{%s}%s"%(namespace,"body"):
                self.body = child[0]
    def write(self):
        self.content.write(self._arquivo, pretty_print=False, encoding="utf-8", method="xml", xml_declaration=True)

class Styles:
    def __init__(self, arquivo=""):
        self._styles = None

        self.styles_doc_all = None
        self.font = None
        self.styles = None
        self.page = None
        self.header = None
        self.header_first = None
        if arquivo:
            self.read(arquivo)

    def read(self, arquivo):
        self._arquivo = arquivo
        self._styles = etree.parse(arquivo)
        for child in self._styles.getroot():
            if child.tag == "{%s}%s"%(namespace,"styles"):
                self.styles_doc_all = child
            elif child.tag == "{%s}%s"%(namespace,"font-face-decls"):
                self.font = child
            elif child.tag == "{%s}%s"%(namespace,"automatic-styles"):
                self.styles = child
            elif child.tag == "{%s}%s"%(namespace,"master-styles"):
                self.page = child[0]
                for h in self.page:
                    if h.tag == "{%s}%s"%(NSMAP["style"],"header"):
                        self.header = h
                    elif h.tag == "{%s}%s"%(NSMAP["style"],"header-first"):
                        self.header_first = h

    def write(self):    
        self._styles.write(self._arquivo, pretty_print=False, encoding="utf-8", method="xml", xml_declaration=True)

class ArquivoODT:
    def __init__(self):
        self.nome = ""
        self.content = None
        self.styles  = None
        self._data = {}# a linha inteira o arquivo csv, enviar tudo para o "assunto" .ini
        self._tmpDirURL = ""
        self._result = {}
        self._zipdata = None
        self._configure = None

    def setConfigure(self, configure):
        self._configure = configure

    def extrair(self):
        #
        # Unzip ODT
        #
        tmpDir="~odt_contents"+self.nome
        self._tmpDirURL=self._configure.rootDirURL+"\\tmp\\"+tmpDir
        print (" -- Extraindo ---------------------")
        print ("%s -> %s" % (self._configure.zipSourceFileURL, self._tmpDirURL))

        self._zipdata = zipfile.ZipFile(self._configure.zipSourceFileURL)
        self._zipdata.extractall(self._tmpDirURL)

        # abrimos os arquivos principais...
        
        xmlFile1="content.xml"
        xmlFileURL1=self._tmpDirURL+"\\"+xmlFile1
        self.content = Content(xmlFileURL1)
        
        xmlFile2="styles.xml"
        xmlFileURL2=self._tmpDirURL+"\\"+xmlFile2
        self.styles = Styles(xmlFileURL2)
    def comprimir(self):
        zipOutFile=self.nome+".odt"
        zipOutFileURL=self._configure.rootDirURL+"\\prontos\\"+zipOutFile
        try:
            os.mkdir(self._configure.rootDirURL+"\\prontos")
        except OSError as error:
            print(error)
        print (" -- Comprimindo --------------------")
        print ("%s -> %s" % (self._tmpDirURL , zipOutFileURL))

        with zipfile.ZipFile(zipOutFileURL, 'w') as outzip:
            zipinfos = self._zipdata.infolist()
            for zipinfo in zipinfos:
                fileName=zipinfo.filename # The name and path as stored in the archive
                fileURL=self._tmpDirURL+"\\"+fileName # The actual name and path
                outzip.write(fileURL,fileName)
    def set_data(self, linha, headers=[]):
        i = 0
        for header in headers:
            self._data[header] = linha[i]
            i=i+1
    def get_data(self):
        return self._data

    def run(self):
        self.styles.write()
        self.content.write()
    
