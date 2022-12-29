from datetime import date, datetime
import sys
# from enum import Enum

sys.path.append("..")

MES = ["-", "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
             "JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]

from tools.odt import namespace, NSMAP, _insert_namespace, new_element, Content, Styles
from tools.odt import ArquivoODT as Arquivo

class ArquivoODT(Arquivo):
    def __init__(self):
        super().__init__()

    def _taf(self, section):
        res = []
        table = new_element("table", attrs=[("name",section),("style-name","Tabela3")], nsmap="table")
        # Construimos as columas....
        table.append( new_element("table-column", attrs=[("style-name","Tabela3.A")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","Tabela3.B")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","Tabela3.C")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","Tabela3.D")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","Tabela3.E")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","Tabela3.F")], nsmap="table") )
        # Construimos a primeira linha como cabeçalho...
        header = new_element("table-row", attrs=[("style-name","Tabela3.1")], nsmap="table")
        body = new_element("table-row", attrs=[("style-name","Tabela3.1")], nsmap="table")
        # inserimos o cabeçalho
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P75")], text="GRAD") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P75")], text="NOME") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P75")], text="MENÇÃO") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P75")], text="PBD") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P75")], text="PAD") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P75")], text="Obs") )
        header.append( cell )
        
        table.append( header )

        # inserimos o Corpo
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P85")], text=self._data.get("grad","Sd Ef Vrv")) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P85")], text=self._data["nome"]) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P85")], text=self._data.get("%s_MEN"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P85")], text=self._data.get("%s_PBD"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P85")], text=self._data.get("%s_PAD"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","Tabela3.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","P85")], text=self._data.get("%s_OBS"%section, " - ") ) )
        body.append( cell )

        table.append( body )
        
        res.append(table)
        return res

    def populate(self):
        self._result = {
            "cabecalho":None,
            "parte_1":[],
            "parte_2":[],
        }
        blank_line= lambda: new_element("p", attrs=[("style-name","P8")])
        #
        # abrimos o arquivo de configuração
        #
        config = self._configure.config
        # datetime.strptime(str_date, "%d/%m/%Y").date()

        meses = []
        for mes in range(1,13):
            res = []
            mes_el = new_element("p", attrs=[("style-name","P10")])
            mes_el.append( new_element("span", attrs=[("style-name","T37")], text = MES[mes]) )
            create = 0
            for section in config.sections():
                if section == "DEFAULT":
                    continue
                # Verificamos se é do mes
                data_event = datetime.strptime(config[section]["data"], "%d/%m/%Y").date()
                if data_event.month != mes:
                    continue
                # verificamos se na lista tem esta sessão ativada para gerar...
                # print(section+"="+self._data.get(section, ""))
                if section in self._data.keys() and (self._data.get(section) not in ['X','x', 1, "1"]):
                    print("interrompemos o "+section)
                    continue
                if create == 0:
                    create = 1
                    mes_el.append( new_element("span", attrs=[("style-name","T18")], text = ":") )
                    res.append( mes_el )
                # Add Titulo do Assunto
                titulo = new_element("p", attrs=[("style-name","P40")], text = config[section]["titulo"])
                if config[section]["tipo"]:
                    titulo.append( new_element("span", attrs=[("style-name","T15")], text = " – %s"%config[section]["tipo"]) )
                res.append( titulo )
                # Colocamos o Assunto
                assunto = new_element("p", attrs=[("style-name","P39")])
                assunto.append( new_element("span", attrs=[("style-name","T50")], text = "-a %d,"%data_event.day) )
                assunto.append( new_element("span", attrs=[("style-name","T50")], text = " %s Nº %d - "%(config[section]["doc_tipo"],int(config[section]["doc_numero"]) )) )
                assunto.append( new_element("span", attrs=[("style-name","T50")], text = config[section]["assunto"].format(**self._data)) )
                res.append( assunto )
                if config[section].get("modo", "normal") == "taf":
                    #add tabela TAF
                    res.extend(self._taf(section))# config[section])
                res.append( blank_line() )
            if not res:
                mes_el.append( new_element("span", attrs=[("style-name","T18")], text = ": Sem alteração.") )
                res.append( mes_el )
                res.append( blank_line() )
            meses.append(res)
        #
        # 1° Parte
        #
        self._result["parte_1"].append( blank_line() )
        self._result["parte_1"].append( new_element("p", attrs=[("style-name","P20")], text = "1ª PARTE:") )
        self._result["parte_1"].append( blank_line() )

        for mes in meses:
            for linha in mes:
                self._result["parte_1"].append(linha)
        self._result["parte_1"].append( blank_line() )
        self._result["parte_1"].append( new_element("p", attrs=[("style-name","P58")], text = "Comportamento: “BOM”") )
        #
        # 2° Parte
        #
        self._result["parte_2"].append( blank_line() )
        self._result["parte_2"].append( new_element("p", attrs=[("style-name","P20")], text = "2ª PARTE:") )
        self._result["parte_2"].append( blank_line() )
        def get_tempo(titulo, tempo, sequencia=0):
            assunto = new_element("p", attrs=[("style-name","P18")])
            assunto.append( new_element("span", attrs=[("style-name","T22")], text=titulo) )# .ljust(101, ".")
            assunto.append( new_element("span", attrs=[("style-name","T22")], text="".ljust(sequencia, ".") ) )
            assunto.append( new_element("span", attrs=[("style-name","T22")], text=tempo) )
            return assunto
        self._result["parte_2"].append( get_tempo("1. TEMPO COMPUTADO DE EFETIVO SERVIÇO (TC)", "00a10m00d", sequencia=58) )
        self._result["parte_2"].append( get_tempo("a. Arregimentado", "00a10m00d", sequencia=111) )
        self._result["parte_2"].append( get_tempo("b. Não Arregimentado", "00a10m00d", sequencia=103) )
        self._result["parte_2"].append( get_tempo("2. TEMPO NÃO COMPUTADO (TNC)", "00a10m00d", sequencia=82) )
        self._result["parte_2"].append( get_tempo("3. TEMPO DE SERVIÇO COMPUTÁVEL PARA MEDALHA MILITAR", "00a10m00d", sequencia=38) )
        self._result["parte_2"].append( get_tempo("4. TEMPO DE SERVIÇO NACIONAL RELEVANTE (TSNR)", "00a10m00d", sequencia=53) )
        self._result["parte_2"].append( get_tempo("5. TEMPO TOTAL DE EFETIVO DE SERVIÇO (TTES)", "00a10m00d", sequencia=58+3) )

        # info do local e data
        self._result["parte_2"].append( blank_line() )
        info = new_element("p", attrs=[("style-name","P19")])
        info.append( new_element("span", attrs=[("style-name","T25")], text="Quartel em %s, %s"%(config["DEFAULT"]["cidade_uf"], config["DEFAULT"]["data"])) )
        self._result["parte_2"].append( info )
        self._result["parte_2"].append( blank_line() )

        #assina...
        self._result["parte_2"].append( blank_line() )
        self._result["parte_2"].append( new_element("p", attrs=[("style-name","P108")], text = config["DEFAULT"]["assina"]) )
        self._result["parte_2"].append( new_element("p", attrs=[("style-name","P108")], text = config["DEFAULT"]["assina_func"]) )
        # self._result["parte_2"].append( blank_line() )

    def run_style(self):
        #
        # abrimos o arquivo de configuração
        #
        config = self._configure.config
        
        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="44ºBIMTZ ") )
        p.append( new_element("s", attrs=[("c","131")]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text="FOLHA Nº 0") )
        p.append( new_element("page-number", attrs=[("select-page","current")], text="9") )
        self.styles.header.append( p )
        self.styles.header.append( new_element("p", attrs=[("style-name","MP1")]) )
        
        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="Continuação das Folhas de Alteração ") )
        p.append( new_element("s", attrs=[("c","55")]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=config["DEFAULT"]["semestre"]) )
        self.styles.header.append( p )

        p = new_element("p", attrs=[("style-name","MP3")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="do %s "%self._data.get("grad","Sd Ef Vrv")) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=self._data["nome"]) )
        len_p = 75-len( self._data.get("grad","Sd Ef Vrv") )-len(self._data["nome"])
        p.append( new_element("s", attrs=[("c",str(len_p) )]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text="PERÍODO: %s"%config["DEFAULT"]["periodo"]) )
        self.styles.header.append( p )

        # Primeira Pagina...
        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="MINISTÉRIO DA DEFESA ") )
        p.append( new_element("s", attrs=[("c","79")]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text="FOLHA DE ALTERAÇÕES") )
        self.styles.header_first.append( p )

        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="EXÉRCITO BRASILEIRO ") )
        p.append( new_element("s", attrs=[("c","82")]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=config["DEFAULT"]["guarnicao"]) )
        self.styles.header_first.append( p )

        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="%s - %s"%(config["DEFAULT"]["comando_militar"],config["DEFAULT"]["regiao_militar"])) )
        p.append( new_element("s", attrs=[("c","99")]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text="FOLHA Nº 01") )
        self.styles.header_first.append( p )

        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text=config["DEFAULT"]["unidade"]) )
        self.styles.header_first.append( p )

        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="NOME: ") )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=self._data["nome"]) )
        self.styles.header_first.append( p )

        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="POSTO/GRADUAÇÃO: ") )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=self._data.get("grad","Sd Ef Vrv")) )
        self.styles.header_first.append( p )

        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="ARMA/SERVIÇO/QUALIFICAÇÃO: ") )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=self._data.get("qm","RECRUTA")) )
        p.append( new_element("s", attrs=[("c", str(53-len(self._data.get("qm","RECRUTA")))  )]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=config["DEFAULT"]["semestre"]) )
        self.styles.header_first.append( p )

        p = new_element("p", attrs=[("style-name","MP1")])
        p.append( new_element("span", attrs=[("style-name","MT5")], text="Identidade Militar: ") )
        p.append( new_element("span", attrs=[("style-name","MT5")], text=self._data["identidade"]) )
        p.append( new_element("s", attrs=[("c", str(74-len(self._data["identidade"])) )]) )
        p.append( new_element("span", attrs=[("style-name","MT5")], text="PERÍODO: %s"%config["DEFAULT"]["periodo"]) )
        self.styles.header_first.append( p )
        
        self.styles.write()

    def run(self):
        self.run_style()
        if not self._result:
            self.populate()
        for p1 in self._result["parte_1"]:
            self.content.body.append(p1)
        for p2 in self._result["parte_2"]:
            self.content.body.append(p2)
        self.content.write()


if __name__ == "__main__":
    from tools.configure import Configure
    from os import path

    # definimos a raiz do diretorio
    rootDirURL = path.dirname(__file__)
    configure = Configure(rootDirURL)
    # definimos o modelo Alteração
    configure.setzipSourceFile("modelo_alteracao.odt")
    # definimos o arquivo lista
    configure.setcsvSourceFile("lista.csv")
    # definimos o arquivo de configuração
    configure.setiniSourceFile("configure.ini")

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

