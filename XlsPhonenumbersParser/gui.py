# coding: utf-8

import tkinter as tk
from tkinter import *
from tkinter.messagebox import *
from tkinter.filedialog import *
import subprocess as sub
from tkinter import font
from controller import XlsPhonenumbersParser


class XlsPhonenumbersParserGUI():
    def __init__(self, root):
        self.root = root
        self.frame=Frame(root)
        self.frame.configure(padx=20, pady=10)
        
        self.controller = XlsPhonenumbersParser()
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
            font=self.customFont,
            command=lambda: self.add_file()
        )
        self.add_file_button.grid(row=0, column=0)

        self.analyze_button = tk.Button(
            self.frame,
            text='PARSE FILES',
            width=30,
            font=self.customFont,
            command=lambda: self.parse_files()
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
            text='Files to parse',
            font=self.customFont
        )
        self.label1.grid(row=2, column=0, sticky=W)

        self.label2 = tk.Label(
            self.frame,
            text='Files parsed',
            font=self.customFont
        )
        self.label2.grid(row=4, column=0, sticky=W)

    def init_files_field(self):
        self.files_field = tk.Text(
            self.frame,
            bg='grey',
            padx=10,
            pady=10,
            font=self.customFont,
            width=50,
            height=5
        )
        self.files_field.grid(row=3, column=0, columnspan=2, sticky=W+E)

    def init_output(self):
        self.out = tk.Text(
            self.frame,
            bg='grey',
            padx=10,
            pady=10,
            font=self.customFont,
            width=50,
            height=15
        )
        self.out.grid(row=5, column=0, columnspan=2, sticky=W+E)

    def add_file(self):
        files = askopenfilenames(
            title='Choose a file to parse',
            filetypes=[
                ('xls files', '.xls'),
                ('xlsx files', '.xlsx'),
                ('ods files', '.ods'),
                ('all files', '.*')
            ]
        )
        if files:
            for f in files:
                self.controller.add_file(f)
                self.files_field.insert('end', '{}\n'.format(f))

    def clear_files(self):
        del self.controller.sheets[:]
        del self.controller.new_sheets[:]
        self.out.delete(1.0, 'end')
        self.files_field.delete(1.0, 'end')

    def close_windows(self):
        """ Close an app window """
        self.root.destroy()

    def parse_files(self):
        self.out.delete(1.0, 'end')
        result = self.controller.parse()
        modified_numbers_list = []

        for sheet in self.controller.new_sheets:
            self.out.insert('end', '{}\n'.format(sheet['name']))
            for numbers in sheet['numbers_modified']:
                if numbers['old'] != numbers['new'] and numbers['old'] not in modified_numbers_list:
                    self.out.insert('end', '*** {} ===> {}\n'.format(numbers['old'], numbers['new']))
                    modified_numbers_list.append(numbers['old'])





