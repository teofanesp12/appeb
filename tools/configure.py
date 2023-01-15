import configparser
import tablib
from datetime import datetime
from os import path

class Configure:
    def __init__(self, root):
        self.rootDirURL = root
        self._data = []
        self._sections_ini = []
        self.prontos = path.join(root, "prontos")

    def setzipSourceFile(self, *zipSourceFile):
        self.zipSourceFile = zipSourceFile
        self.zipSourceFileURL=path.join(self.rootDirURL, *zipSourceFile)

    def setcsvSourceFile(self, *csvSourceFile):
        self.csvSourceFile = csvSourceFile
        self.csvSourceFileURL=path.join(self.rootDirURL, *csvSourceFile)

        self._data = tablib.Dataset().load(open(self.csvSourceFileURL, 'rb').read().decode('utf-8'), format='csv')

    def setiniSourceFile(self, *iniSourceFile):
        self.iniSourceFile = iniSourceFile
        self.iniSourceFileURL=path.join(self.rootDirURL, *iniSourceFile)
        self.config = configparser.ConfigParser()
        self.config.read(self.iniSourceFileURL, encoding='utf-8')

    def getFileStr(self, dirname, filename):
        return dirname+"\\"+filename

    def getData(self):
        return self._data

    def getSections(self):
        if self._sections_ini:
            return self._sections_ini
        sections = self.config.sections().copy()
        # print (sections)
        meses = {}
        for i in range(1,13):
            meses[i] = []
            for e in range(1,32):
                for section in sections.copy():
                    data_event = datetime.strptime(self.config[section]["data"], "%d/%m/%Y").date()
                    if data_event.day == e and data_event.month == i:
                        meses[i].append(section)
                        try:
                            sections.remove(section)
                        except ValueError:
                            print("Não consegui resolver "+section)
        # print(meses)
        for i in range(1,13):
            for sec in meses[i]:
                self._sections_ini.append(sec)
        # print (self._sections_ini)
        return self._sections_ini
