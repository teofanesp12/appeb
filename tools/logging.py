try:
    from tk import NORMAL, INSERT
except:
    INSERT = NORMAL = None

from threading import Thread
class AsyncConsole(Thread):
    def __init__(self, url):
        super().__init__()

        self.html = None
        self.url = url

    def run(self):
        response = requests.get(self.url)
        self.html = response.text

def info(tk, txt):
    if tk:
        #tk.config(state = NORMAL)
        #if txt:
        #    tk.insert(INSERT, txt)
        #tk.insert(INSERT, " \n")
        #tk.config(state = DISABLED)
        print(txt)
    else:
        print(txt)
