# coding: utf-8

import tkinter as tk
from tkinter.messagebox import *
from tkinter.filedialog import *


class XlsDuplicateParserGUI():
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.menubar = Menu(self.master)
        self.create_menu()
        self.quitButton = tk.Button(
            self.frame,
            text='Quit',
            width=25,
            command=self.close_windows
        )
        self.quitButton.grid(row=10, column=0)
        self.master.config(menu=self.menubar)
        self.frame.pack()

    def create_menu(self):
        self.menu1 = Menu(self.menubar, tearoff=0)
        self.menu1.add_command(
            label="Importer fichier 1",
            command=lambda: self.select_file(1)
        )
        self.menu1.add_command(
            label="Importer fichier 2",
            command=lambda: self.select_file(2)
        )
        self.menu1.add_separator()
        self.menu1.add_command(label="Quitter", command=self.master.quit)
        self.menubar.add_cascade(label="Fichier", menu=self.menu1)

    def select_file(self, num):
        filename = askopenfilename(
            title="Ouvrir un fichier",
            filetypes=[
                ('xls files', '.xls'),
                ('xlsx files', '.xlsx'),
                ('all files', '.*')
            ]
        )

    def close_windows(self):
        self.master.destroy()
