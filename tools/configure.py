import configparser
import tablib

class Configure:
    def __init__(self, root):
        self.rootDirURL = root
        self._data = []

    def setzipSourceFile(self, zipSourceFile):
        self.zipSourceFile = zipSourceFile
        self.zipSourceFileURL=self.rootDirURL+"\\"+zipSourceFile

    def setcsvSourceFile(self, csvSourceFile):
        self.csvSourceFile = csvSourceFile
        self.csvSourceFileURL=self.rootDirURL+"\\"+csvSourceFile

        self._data = tablib.Dataset().load(open(self.csvSourceFileURL, 'r'))

    def setiniSourceFile(self, iniSourceFile):
        self.iniSourceFile = iniSourceFile
        self.iniSourceFileURL=self.rootDirURL+"\\"+iniSourceFile
        self.config = configparser.ConfigParser()
        self.config.read(self.iniSourceFileURL, encoding='utf-8')

    def getFileStr(self, dirname, filename):
        return dirname+"\\"+filename

    def getData(self):
        return self._data
