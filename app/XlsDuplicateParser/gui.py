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
        self.init_action_buttons()
        self.init_output()

        self.row_where_add_file = 5

        self.frame.grid()

    def add_file(self):
        file_name = askopenfilename(
            title='Choose a file to compare',
            filetypes=[
                ('xls files', '.xls'),
                ('xlsx files', '.xlsx'),
                ('all files', '.*')
            ]
        )
        if file_name:
            self.controller.add_sheet_to_compare(file_name)
    
            files_label = tk.Label(
                self.frame,
                text="File to analyze : {}".format(file_name)
            )
            files_label.grid(row=self.row_where_add_file, column=0, sticky=W+E)
    
            self.row_where_add_file += 1

    def clear_files(self):
        del self.controller.sheets_to_compare[:]
        self.out.delete(0, 'end')

    def init_menu(self):
        self.menubar = Menu(self.master)
        self.menu1 = Menu(self.menubar, tearoff=0)
        self.menu1.add_command(
            label="Import file to compare",
            command=lambda: self.add_file()
        )
        self.menu1.add_separator()
        self.menu1.add_command(label="Quitter", command=self.master.quit)
        self.menubar.add_cascade(label="Fichier", menu=self.menu1)
        self.master.config(menu=self.menubar)

    def init_action_buttons(self):
        self.add_file_button = tk.Button(
            self.frame,
            text='ADD FILE',
            width=30,
            command=lambda: self.add_file()
        )
        self.add_file_button.grid(row=0, column=0, columnspan=6)

        self.analyze_button = tk.Button(
            self.frame,
            text='FIND OCCURENCES',
            width=30,
            command=lambda: self.analyze()
        )
        self.analyze_button.grid(row=1, column=0, columnspan=6)

        self.add_file_button = tk.Button(
            self.frame,
            text='CLEAR FILES',
            width=30,
            command=lambda: self.clear_files()
        )
        self.add_file_button.grid(row=2, column=0, columnspan=6)

        self.close_button = tk.Button(
            self.frame,
            text='QUIT',
            width=30,
            command=lambda: self.close_windows()
        )
        self.close_button.grid(row=3, column=0, columnspan=6)

    def init_output(self):
        self.out = tk.Text(self.frame)
        self.out.grid(row=0, rowspan=13, column=8, columnspan=6)

    def close_windows(self):
        """ Close an app window """
        self.master.destroy()

    def analyze(self):
        result = self.controller.browse_rows_multiple_files(self.controller.sheets_to_compare)
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
