# !/usr/bin/env python3
# -*- coding: utf-8 -*-
from tools.configure import Configure
from tools.system import libreoffice_write, libreoffice_calc, editortxt, explore

from os import path
from assentamentos.app import ArquivoODT

# definimos a raiz do diretorio
rootDirURL = path.dirname(__file__)

import os
from threading import Thread

import tkinter
import tkinter.messagebox
import customtkinter

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()

        self.lista_csv = os.path.join(rootDirURL, "assentamentos","lista.csv")
        self.configure_ini = os.path.join(rootDirURL, "assentamentos","configure.ini")

        # configure window
        self.title("Alterações e Assentamentos")
        # self.iconphoto(False, tkinter.PhotoImage(file = os.path.join(rootDirURL, "static", "militar.png")))
        self.geometry(f"{1100}x{580}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=6, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="AÇÕES", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, command=self.open_prontos, text="Abrir PRONTOS")
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, command=self.run_def, text="Gerar Arquivos")
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.conf_label = customtkinter.CTkLabel(self.sidebar_frame, text="CONFIGURAÇÕES", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.conf_label.grid(row=3, column=0, padx=20, pady=(10, 10))
        self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, command=self.open_ini, text="Arquivo CONF")
        self.sidebar_button_3.grid(row=4, column=0, padx=20, pady=10)
        self.sidebar_button_4 = customtkinter.CTkButton(self.sidebar_frame, command=self.open_csv, text="Planilha")
        self.sidebar_button_4.grid(row=5, column=0, padx=20, pady=10)
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Appearance Mode:", anchor="w")
        self.appearance_mode_label.grid(row=8, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=9, column=0, padx=20, pady=(10, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=10, column=0, padx=20, pady=(10, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                               command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=11, column=0, padx=20, pady=(10, 20))

        # create main entry and button
        self.entry_ini = customtkinter.CTkEntry(self, placeholder_text="INI")
        self.entry_ini.grid(row=3, column=1, columnspan=2, padx=(20, 0), pady=(20, 10), sticky="nsew")
        self.entry_ini.insert(tkinter.END, self.configure_ini)

        self.main_button_1 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), command=self.select_file_ini, text='Trocar')
        self.main_button_1.grid(row=3, column=3, padx=(20, 20), pady=(20, 10), sticky="nsew")

        self.entry_csv = customtkinter.CTkEntry(self, placeholder_text="CSV")
        self.entry_csv.grid(row=4, column=1, columnspan=2, padx=(20, 0), pady=(10, 10), sticky="nsew")
        self.entry_csv.insert(tkinter.END, self.lista_csv)

        self.main_button_2 = customtkinter.CTkButton(master=self, fg_color="transparent", border_width=2, text_color=("gray10", "#DCE4EE"), command=self.select_file_csv, text='Trocar')
        self.main_button_2.grid(row=4, column=3, padx=(20, 20), pady=(10, 10), sticky="nsew")

        # create textbox
        self.info = customtkinter.CTkTextbox(self, width=250)
        self.info.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.info.configure(state=customtkinter.DISABLED)

        # create tabview
        self.tabview = customtkinter.CTkTabview(self, width=250)
        self.tabview.grid(row=0, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.tabview.add("CTkTabview")
        self.tabview.add("Tab 2")
        self.tabview.add("Tab 3")
        self.tabview.tab("CTkTabview").grid_columnconfigure(0, weight=1)  # configure grid of individual tabs
        self.tabview.tab("Tab 2").grid_columnconfigure(0, weight=1)

        self.optionmenu_1 = customtkinter.CTkOptionMenu(self.tabview.tab("CTkTabview"), dynamic_resizing=False,
                                                        values=["Value 1", "Value 2", "Value Long Long Long"])
        self.optionmenu_1.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.combobox_1 = customtkinter.CTkComboBox(self.tabview.tab("CTkTabview"),
                                                    values=["Value 1", "Value 2", "Value Long....."])
        self.combobox_1.grid(row=1, column=0, padx=20, pady=(10, 10))
        self.string_input_button = customtkinter.CTkButton(self.tabview.tab("CTkTabview"), text="Open CTkInputDialog",
                                                           command=self.open_input_dialog_event)
        self.string_input_button.grid(row=2, column=0, padx=20, pady=(10, 10))
        self.label_tab_2 = customtkinter.CTkLabel(self.tabview.tab("Tab 2"), text="CTkLabel on Tab 2")
        self.label_tab_2.grid(row=0, column=0, padx=20, pady=20)

        # create radiobutton frame
        self.radiobutton_frame = customtkinter.CTkFrame(self)
        self.radiobutton_frame.grid(row=0, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.radio_var = tkinter.IntVar(value=0)
        self.label_radio_group = customtkinter.CTkLabel(master=self.radiobutton_frame, text="CTkRadioButton Group:")
        self.label_radio_group.grid(row=0, column=2, columnspan=1, padx=10, pady=10, sticky="")
        self.radio_button_1 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=0)
        self.radio_button_1.grid(row=1, column=2, pady=10, padx=20, sticky="n")
        self.radio_button_2 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=1)
        self.radio_button_2.grid(row=2, column=2, pady=10, padx=20, sticky="n")
        self.radio_button_3 = customtkinter.CTkRadioButton(master=self.radiobutton_frame, variable=self.radio_var, value=2)
        self.radio_button_3.grid(row=3, column=2, pady=10, padx=20, sticky="n")

        # create slider and progressbar frame
        self.slider_progressbar_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        self.slider_progressbar_frame.grid(row=1, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.slider_progressbar_frame.grid_columnconfigure(0, weight=1)
        self.slider_progressbar_frame.grid_rowconfigure(4, weight=1)
        #self.seg_button_1 = customtkinter.CTkSegmentedButton(self.slider_progressbar_frame)
        #self.seg_button_1.grid(row=0, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
        self.progressbar_1 = customtkinter.CTkProgressBar(self.slider_progressbar_frame)
        self.progressbar_1.grid(row=0, column=0, padx=(20, 10), pady=(10, 10), sticky="ew")
        self.progressbar_1.set(0)
        # create textbox
        self.info2 = customtkinter.CTkTextbox(self.slider_progressbar_frame, width=250)
        self.info2.grid(row=1, column=0, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.info2.configure(state=customtkinter.DISABLED)
        

        # create scrollable frame
        self.scrollable_frame = customtkinter.CTkScrollableFrame(self, label_text="CTkScrollableFrame")
        self.scrollable_frame.grid(row=1, column=2, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
        self.scrollable_frame_switches = []
        for i in range(15):
            switch = customtkinter.CTkSwitch(master=self.scrollable_frame, text=f"CTkSwitch {i}")
            switch.grid(row=i, column=0, padx=10, pady=(0, 20))
            self.scrollable_frame_switches.append(switch)

        # create checkbox and switch frame
        self.checkbox_slider_frame = customtkinter.CTkFrame(self)
        self.checkbox_slider_frame.grid(row=1, column=3, padx=(20, 20), pady=(20, 0), sticky="nsew")
        self.checkbox_1 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        self.checkbox_1.grid(row=1, column=0, pady=(20, 0), padx=20, sticky="n")
        self.checkbox_2 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        self.checkbox_2.grid(row=2, column=0, pady=(20, 0), padx=20, sticky="n")
        self.checkbox_3 = customtkinter.CTkCheckBox(master=self.checkbox_slider_frame)
        self.checkbox_3.grid(row=3, column=0, pady=20, padx=20, sticky="n")

        # set default values
        # self.sidebar_button_3.configure(state="disabled", text="Disabled CTkButton")
        self.checkbox_3.configure(state="disabled")
        self.checkbox_1.select()
        self.scrollable_frame_switches[0].select()
        self.scrollable_frame_switches[4].select()
        self.radio_button_3.configure(state="disabled")
        self.appearance_mode_optionemenu.set("Dark")
        self.scaling_optionemenu.set("100%")
        self.optionmenu_1.set("CTkOptionmenu")
        self.combobox_1.set("CTkComboBox")
        #self.slider_1.configure(command=self.progressbar_2.set)
        #self.slider_2.configure(command=self.progressbar_3.set)
        #self.progressbar_1.configure(mode="indeterminnate")
        #self.progressbar_1.start()
        # self.info.insert("0.0", "CTkTextbox\n\n" + "Lorem ipsum dolor sit amet, consetetur sadipscing elitr, sed diam nonumy eirmod tempor invidunt ut labore et dolore magna aliquyam erat, sed diam voluptua.\n\n" * 20)
        #self.seg_button_1.configure(values=["CTkSegmentedButton", "Value 2", "Value 3"])
        #self.seg_button_1.set("Value 2")

    def open_input_dialog_event(self):
        dialog = customtkinter.CTkInputDialog(text="Type in a number:", title="CTkInputDialog")
        print("CTkInputDialog:", dialog.get_input())

    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def open_csv(self):
        listaURL = os.path.join(rootDirURL, "assentamentos", "lista.csv")
        libreoffice_calc(listaURL)
    def open_prontos(self):
        prontosURL = os.path.join(rootDirURL, "prontos")
        try:
            self.info.insert(customtkinter.INSERT, "criando arquivo...\n")
            os.mkdir(prontosURL)
        except OSError as error:
            self.info.configure(state=customtkinter.NORMAL)
            self.info.insert(customtkinter.INSERT, "arquivo já criado...\n")
            self.info.configure(state=customtkinter.DISABLED)
            # print("arquivo já criado...")
        Thread(target=explore, args=(prontosURL,)).start()
    def open_ini(self):
        configureURL = os.path.join(rootDirURL, "assentamentos", "configure.ini")
        editortxt(configureURL)

    def run_def(self):
        Thread(target=self.run).start()
    def run(self):
        configure = Configure(rootDirURL)
        # definimos o modelo Alteração
        configure.setzipSourceFile("assentamentos","modelo_alteracao.odt")
        # definimos o arquivo lista
        configure.setcsvSourceFile(self.lista_csv)
        # definimos o arquivo de configuração
        configure.setiniSourceFile(self.configure_ini)
        #
        # iniciamos...
        #
        self.info.configure(state=customtkinter.NORMAL)
        self.info2.configure(state=customtkinter.NORMAL)
        self.info.insert(customtkinter.INSERT, "Gerando Documentos ")
        len_data = len(configure.getData())
        self.progressbar_1.set(0)
        i = 0
        for linha in configure.getData():
            i+=1
            self.info2.insert(customtkinter.INSERT, linha[0])
            arquivo = ArquivoODT()
            arquivo.nome = linha[0]
            arquivo.setConfigure(configure)
            arquivo.set_data(linha, headers=configure.getData().headers)
            # self.info.insert(customtkinter.INSERT, ": Iniciou com sucesso")
            arquivo.extrair()
            arquivo.run()
            arquivo.comprimir()
            self.info2.insert(customtkinter.INSERT, " - arquivo ODT pronto com sucesso\n\n")
            self.progressbar_1.set(i/len_data)
        self.info.insert(customtkinter.INSERT, "- PRONTO(S)\n")
        self.info.configure(state=customtkinter.DISABLED)
        self.info2.configure(state=customtkinter.DISABLED)
        self.progressbar_1.set(1)

    def select_file_ini(self):
        filetypes = (
            ('Arquivo CONF','*.ini'),
        )
        filename = tkinter.filedialog.askopenfilename(
            title='Abrir Arquivo para Configuração',
            initialdir=os.path.join(rootDirURL, "assentamentos"),
            filetypes=filetypes
        )
        # tkinter.messagebox.showinfo(title='teste', message=filename)
        self.entry_ini.delete(0,tkinter.END)
        self.entry_ini.insert(tkinter.END, filename)
        self.configure_ini = filename
    def select_file_csv(self):
        filetypes = (
            ('Arquivo Planilha','*.csv'),
        )
        filename = tkinter.filedialog.askopenfilename(
            title='Abrir Arquivo Planilha',
            initialdir=os.path.join(rootDirURL, "assentamentos"),
            filetypes=filetypes
        )
        # tkinter.messagebox.showinfo(title='teste', message=filename)
        self.entry_csv.delete(0,tkinter.END)
        self.entry_csv.insert(tkinter.END, filename)
        self.lista_csv = filename


if __name__ == "__main__":
    app = App()
    app.mainloop()
