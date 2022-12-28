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
