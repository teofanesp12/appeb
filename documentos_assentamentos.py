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
