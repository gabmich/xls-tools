# coding: utf-8

import tkinter as tk
from tkinter import *
from tkinter.messagebox import *
from tkinter.filedialog import *
import subprocess as sub
from tkinter import font
from controller import XlsDuplicateParser


class XlsDuplicateParserGUI():
    def __init__(self, root):
        self.root = root
        self.frame=Frame(root)
        self.frame.configure(padx=20, pady=10)

        self.controller = XlsDuplicateParser()
        self.customFont = font.Font(family="Helvetica", size=20)

        self.init_menu()
        self.init_action_buttons()
        self.init_labels()
        self.init_files_field()
        self.init_output()

        self.frame.grid()

        for row_index in range(0, 6):
            Grid.rowconfigure(self.frame, row_index, weight=1)

        for col_index in range(20):
            Grid.columnconfigure(self.frame, col_index, weight=1)

    def init_menu(self):
        self.menubar = Menu(self.root)
        self.menu1 = Menu(self.menubar, tearoff=0)
        self.menu1.add_command(
            label="Add file(s) to compare",
            command=lambda: self.add_file()
        )
        self.menu1.add_separator()
        self.menu1.add_command(label="Quitter", command=self.root.quit)
        self.menubar.add_cascade(label="Fichier", menu=self.menu1)
        self.root.config(menu=self.menubar)

    def init_action_buttons(self):
        self.add_file_button = tk.Button(
            self.frame,
            text='ADD FILE(S)',
            width=30,
            bg='blue',
            font=self.customFont,
            command=lambda: self.add_file()
        )
        self.add_file_button.grid(row=0, column=0)

        self.analyze_button = tk.Button(
            self.frame,
            text='FIND OCCURENCES',
            width=30,
            font=self.customFont,
            command=lambda: self.analyze()
        )
        self.analyze_button.grid(row=1, column=0)

        self.clear_files_button = tk.Button(
            self.frame,
            text='CLEAR FILES',
            width=30,
            font=self.customFont,
            command=lambda: self.clear_files()
        )
        self.clear_files_button.grid(row=0, column=1)

        self.close_button = tk.Button(
            self.frame,
            text='QUIT',
            width=30,
            font=self.customFont,
            command=lambda: self.close_windows()
        )
        self.close_button.grid(row=1, column=1)

    def init_labels(self):
        self.label1 = tk.Label(
            self.frame,
            text='Files selected',
            font=self.customFont,
            height=1
        )
        self.label1.grid(row=2, column=0, sticky=W)

        self.label2 = tk.Label(
            self.frame,
            text='Occurences found',
            font=self.customFont,
            height=1
        )
        self.label2.grid(row=4, column=0, sticky=W)

    def init_files_field(self):
        self.files_field = tk.Text(
            self.frame,
            bg='grey',
            padx=10,
            pady=10,
            font=self.customFont,
            height=5
        )
        self.files_field.grid(row=3, column=0, columnspan=2)

    def init_output(self):
        self.out = tk.Text(
            self.frame,
            bg='grey',
            padx=10,
            pady=10,
            font=self.customFont
        )
        self.out.configure(height=10)
        self.out.grid(row=5, column=0, columnspan=2)

    def add_file(self):
        files = askopenfilenames(
            title='Choose a file to compare',
            filetypes=[
                ('xls files', '.xls'),
                ('xlsx files', '.xlsx'),
                ('ods files', '.ods'),
                ('all files', '.*')
            ]
        )
        if files:
            for f in files:
                self.controller.add_sheet_to_compare(f)
                self.files_field.insert('end', '{}\n'.format(f))

    def clear_files(self):
        del self.controller.sheets_to_compare[:]
        self.out.delete(1.0, 'end')
        self.files_field.delete(1.0, 'end')

    def close_windows(self):
        """ Close an app window """
        self.root.destroy()

    def analyze(self):
        self.out.delete(1.0, 'end')
        result = self.controller.browse_rows_multiple_files(
            self.controller.sheets_to_compare
        )
        for r in result:
            self.out.insert('end', 'MATCH : {}\n'.format(r))