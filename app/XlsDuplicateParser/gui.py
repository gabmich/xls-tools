# coding: utf-8

import tkinter as tk
from tkinter.messagebox import *
from tkinter.filedialog import *
from app.XlsDuplicateParser.controller import XlsDuplicateParser
import subprocess as sub
import pyexcel as p


class XlsDuplicateParserGUI():
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.controller = XlsDuplicateParser()

        self.init_menu()
        self.init_labels()
        self.init_action_buttons()
        self.init_output()

        self.frame.grid()

    def init_menu(self):
        self.menubar = Menu(self.master)
        self.menu1 = Menu(self.menubar, tearoff=0)
        self.menu1.add_command(
            label="Importer fichier 1",
            command=lambda: self.openfiledialog(
                title="Select first file",
                target="first"
            )
        )
        self.menu1.add_command(
            label="Importer fichier 2",
            command=lambda: self.openfiledialog(
                title="Select second file",
                target="second"
            )
        )
        self.menu1.add_separator()
        self.menu1.add_command(label="Quitter", command=self.master.quit)
        self.menubar.add_cascade(label="Fichier", menu=self.menu1)
        self.master.config(menu=self.menubar)

    def init_labels(self):
        self.label_for_first_file = tk.Label(
            self.frame,
            text="First file : "
        ).grid(row=0, column=0, sticky=W)

        self.file_name1 = Label(self.frame, text='(use menu to load file)')
        self.file_name1.grid(row=0, column=2, sticky=W)

        self.label_for_second_file = tk.Label(
            self.frame,
            text="Second file : "
        ).grid(row=1, column=0, sticky=W)

        self.file_name2 = Label(self.frame, text='(use menu to load file)')
        self.file_name2.grid(row=1, column=2, sticky=W)

    def init_action_buttons(self):
        self.analyze_button = tk.Button(
            self.frame,
            text='FIND OCCURENCES',
            width=30,
            command=lambda: self.analyze()
        )
        self.analyze_button.grid(row=2, column=1)

        self.close_button = tk.Button(
            self.frame,
            text='QUIT',
            width=30,
            command=lambda: self.close_windows()
        )
        self.close_button.grid(row=3, column=1)

    def init_output(self):
        self.out = tk.Text(self.frame)
        self.out.grid(row=6, columnspan=3)

    def close_windows(self):
        """ Close an app window """
        self.master.destroy()

    def analyze(self):
        result = self.controller.browse_rows()
        for r in result:
            self.out.insert('end', 'MATCH : {}\n'.format(r))

    def openfiledialog(self, title, target):
        filename = askopenfilename(
            title=title,
            filetypes=[
                ('xls files', '.xls'),
                ('xlsx files', '.xlsx'),
                ('all files', '.*')
            ]
        )
        if target == 'first':
            self.file_name1['text'] = filename
            self.controller.sheet1 = self.controller.set_sheet(filename)
        elif target == 'second':
            self.file_name2['text'] = filename
            self.controller.sheet2 = self.controller.set_sheet(filename)

    def close_windows(self):
        self.master.destroy()
