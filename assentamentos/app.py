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
        table = new_element("table", attrs=[("name",section),("style-name","TabelaTAF")], nsmap="table")
        # Construimos as columas....
        table.append( new_element("table-column", attrs=[("style-name","TabelaTAF.A")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","TabelaTAF.B")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","TabelaTAF.C")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","TabelaTAF.D")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","TabelaTAF.E")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","TabelaTAF.F")], nsmap="table") )
        # Construimos a primeira linha como cabeçalho...
        header = new_element("table-row", attrs=[("style-name","TabelaTAF.1")], nsmap="table")
        body = new_element("table-row", attrs=[("style-name","TabelaTAF.1")], nsmap="table")
        # inserimos o cabeçalho
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="GRAD") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="NOME") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="MENÇÃO") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="PBD") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="PAD") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="Obs") )
        header.append( cell )
        
        table.append( header )

        # inserimos o Corpo
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("grad","Sd Ef Vrv")) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data["nome"]) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("%s_MEN"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("%s_PBD"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("%s_PAD"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","TabelaTAF.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("%s_OBS"%section, " - ") ) )
        body.append( cell )

        table.append( body )
        
        res.append(table)
        return res

    def _trans_interna(self, section):
        res = []
        # grad, nome, qmg_qmp, origem, destino, claro, obs
        return res

    def _insp_saude(self, section):
        res = []
        # grad, nome, sessao, data, ata
        table = new_element("table", attrs=[("name",section),("style-name","INSP")], nsmap="table")
        # Construimos as columas....
        table.append( new_element("table-column", attrs=[("style-name","INSP.A")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","INSP.B")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","INSP.C")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","INSP.D")], nsmap="table") )
        table.append( new_element("table-column", attrs=[("style-name","INSP.E")], nsmap="table") )
        # Construimos a primeira linha como cabeçalho...
        header = new_element("table-row", attrs=[("style-name","INSP.1")], nsmap="table")
        body = new_element("table-row", attrs=[("style-name","INSP.2")], nsmap="table")
        # inserimos o cabeçalho
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="GRAD") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="NOME") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="SESSÃO") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="DATA") )
        header.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PHEADERS")], text="ATA") )
        header.append( cell )
        
        table.append( header )

        # inserimos o Corpo
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("grad","Sd Ef Vrv")) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data["nome"]) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("%s_SESSAO"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.D1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("%s_DATA"%section)) )
        body.append( cell )
        
        cell = new_element("table-cell", attrs=[("style-name","INSP.A1"), ['value-type','string','office']], nsmap="table")
        cell.append( new_element("p", attrs=[("style-name","PBODY")], text=self._data.get("%s_ATA"%section)) )
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
        blank_line= lambda: new_element("p", attrs=[("style-name","PBLANK")])
        #
        # abrimos o arquivo de configuração
        #
        config = self._configure.config
        # datetime.strptime(str_date, "%d/%m/%Y").date()

        meses = []
        for mes in range(1,13):
            res = []
            mes_el = new_element("p", attrs=[("style-name","PMES")])
            mes_el.append( new_element("span", attrs=[("style-name","TMES")], text = MES[mes]) )
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
                    mes_el.append( new_element("span", attrs=[("style-name","TMESSUB")], text = ":") )
                    res.append( mes_el )
                # Add Titulo do Assunto
                titulo = new_element("p", attrs=[("style-name","PTITULO")], text = config[section]["titulo"])
                if config[section]["tipo"]:
                    titulo.append( new_element("span", attrs=[("style-name","TTIPO")], text = " – %s"%config[section]["tipo"]) )
                res.append( titulo )
                # Colocamos o Assunto
                assunto = new_element("p", attrs=[("style-name","PASSUNTO")])
                assunto.append( new_element("span", attrs=[("style-name","TASSUNTO")], text = "-a %d,"%data_event.day) )
                assunto.append( new_element("span", attrs=[("style-name","TASSUNTO")], text = " %s Nº %d - "%(config[section]["doc_tipo"],int(config[section]["doc_numero"]) )) )
                i = 0
                for ass in config[section]["assunto"].format(**self._data).split("\\n"):
                    if i == 0:
                        assunto.append( new_element("span", attrs=[("style-name","TASSUNTO")], text = ass) )
                        res.append( assunto )
                    else:
                        p = new_element("p", attrs=[("style-name","PASSUNTO")])
                        # p.append( new_element("s", attrs=[("c", "4" )]) )
                        p.append( new_element("tab" ) )
                        p.append( new_element("span", attrs=[("style-name","TASSUNTO")], text = ass) )
                        res.append(p)
                    i=i+1
                if config[section].get("modo", "normal") == "taf":
                    #add tabela TAF
                    res.extend(self._taf(section))# config[section])
                elif config[section].get("modo", "normal") == "insp_saude":
                    res.extend( self._insp_saude(section) )
                res.append( blank_line() )
            if not res:
                mes_el.append( new_element("span", attrs=[("style-name","TMESSUB")], text = ": Sem alteração.") )
                res.append( mes_el )
                res.append( blank_line() )
            meses.append(res)
        #
        # 1° Parte
        #
        self._result["parte_1"].append( blank_line() )
        self._result["parte_1"].append( new_element("p", attrs=[("style-name","PPARTE")], text = "1ª PARTE:") )
        self._result["parte_1"].append( blank_line() )

        for mes in meses:
            for linha in mes:
                self._result["parte_1"].append(linha)
        self._result["parte_1"].append( blank_line() )
        self._result["parte_1"].append( new_element("p", attrs=[("style-name","PCOMPORTAMENTO")], text = "Comportamento: “%s”"%self._data.get("comportamento","BOM")) )
        #
        # 2° Parte
        #
        self._result["parte_2"].append( blank_line() )
        self._result["parte_2"].append( new_element("p", attrs=[("style-name","PPARTE")], text = "2ª PARTE:") )
        self._result["parte_2"].append( blank_line() )
        def get_tempo(titulo, tempo, sequencia=0):
            assunto = new_element("p", attrs=[("style-name","PTEMPO")])
            assunto.append( new_element("span", attrs=[("style-name","TTEMPO")], text=titulo) )# .ljust(101, ".")
            assunto.append( new_element("span", attrs=[("style-name","TTEMPO")], text="".ljust(sequencia, ".") ) )
            assunto.append( new_element("span", attrs=[("style-name","TTEMPO")], text=tempo) )
            return assunto
        self._result["parte_2"].append( get_tempo("1. TEMPO COMPUTADO DE EFETIVO SERVIÇO (TC)", self._data.get("TCES", "00a10m00d"), sequencia=58) )
        self._result["parte_2"].append( get_tempo("a. Arregimentado", self._data.get("TC_A", "00a10m00d"), sequencia=111) )
        self._result["parte_2"].append( get_tempo("b. Não Arregimentado", self._data.get("TC_NA", "00a10m00d"), sequencia=103) )
        self._result["parte_2"].append( get_tempo("2. TEMPO NÃO COMPUTADO (TNC)", self._data.get("TNC", "00a10m00d"), sequencia=82) )
        self._result["parte_2"].append( get_tempo("3. TEMPO DE SERVIÇO COMPUTÁVEL PARA MEDALHA MILITAR", self._data.get("TSCMM", "01a00m00d"), sequencia=38) )
        self._result["parte_2"].append( get_tempo("4. TEMPO DE SERVIÇO NACIONAL RELEVANTE (TSNR)", self._data.get("TSNR", "00a10m00d"), sequencia=53) )
        self._result["parte_2"].append( get_tempo("5. TEMPO TOTAL DE EFETIVO DE SERVIÇO (TTES)", self._data.get("TTES", "00a10m00d"), sequencia=58+3) )

        # info do local e data
        self._result["parte_2"].append( blank_line() )
        info = new_element("p", attrs=[("style-name","PLOCALDATA")])
        info.append( new_element("span", attrs=[("style-name","TLOCALDATA")], text="Quartel em %s, %s"%(config["DEFAULT"]["cidade_uf"], config["DEFAULT"]["data"])) )
        self._result["parte_2"].append( info )
        self._result["parte_2"].append( blank_line() )

        #assina...
        self._result["parte_2"].append( blank_line() )
        self._result["parte_2"].append( new_element("p", attrs=[("style-name","PASSINA")], text = config["DEFAULT"]["assina"]) )
        self._result["parte_2"].append( new_element("p", attrs=[("style-name","PASSINA")], text = config["DEFAULT"]["assina_func"]) )
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

