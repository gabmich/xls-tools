# coding: utf-8

import tkinter as tk
from tkinter import *
from tkinter.messagebox import *
from tkinter.filedialog import *
import subprocess as sub
from tkinter import font
from controller import PnParser


class PnParserGUI():
    def __init__(self, root):
        self.root = root
        self.frame=Frame(root)
        self.frame.configure(padx=20, pady=10)
        
        self.controller = PnParser()
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
        self.menu1.add_command(label="Quitter", command=self.root.quit)
        self.menubar.add_cascade(label="Fichier", menu=self.menu1)
        self.root.config(menu=self.menubar)

    def init_action_buttons(self):
        self.add_sheet_button = tk.Button(
            self.frame,
            text='IMPORT',
            width=30,
            font=self.customFont,
            command=lambda: self.add_file()
        )
        self.add_sheet_button.grid(row=0, column=0)

        self.clear_sheets_button = tk.Button(
            self.frame,
            text='CLEAR',
            width=30,
            font=self.customFont,
            command=lambda: self.clear_sheets()
        )
        self.clear_sheets_button.grid(row=0, column=1)

        self.parse_button = tk.Button(
            self.frame,
            text='PARSE',
            width=30,
            font=self.customFont,
            command=lambda: self.parse_files()
        )
        self.parse_button.grid(row=1, column=0)

        self.compare_button = tk.Button(
            self.frame,
            text='FIND OCCURENCES',
            width=30,
            font=self.customFont,
            command=lambda: self.compare_files()
        )
        self.compare_button.grid(row=1, column=1)

    def init_labels(self):
        self.label1 = tk.Label(
            self.frame,
            text='Files to parse',
            font=self.customFont
        )
        self.label1.grid(row=3, column=0, sticky=W)

        self.label2 = tk.Label(
            self.frame,
            text='Result',
            font=self.customFont
        )
        self.label2.grid(row=5, column=0, sticky=W)

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
        self.files_field.grid(row=4, column=0, columnspan=2, sticky=W+E)

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
        self.out.grid(row=6, column=0, columnspan=2, sticky=W+E)

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
                self.controller.add_sheet_to_parse(f)
                self.files_field.insert('end', '{}\n'.format(f))

    def clear_sheets(self):
        del self.controller.sheets_to_parse[:]
        del self.controller.sheets_parsed[:]
        self.out.delete(1.0, 'end')
        self.files_field.delete(1.0, 'end')

    def close_windows(self):
        """ Close an app window """
        self.root.destroy()

    def parse_files(self):
        self.out.delete(1.0, 'end')
        result = self.controller.parse()
        modified_numbers_list = []
        changes = open(self.controller.changes, 'w')

        for sheet in self.controller.sheets_parsed:
            self.out.insert('end', '{}\n'.format(sheet['name']))
            changes.write('{}\n'.format(sheet['name']))

            for numbers in sheet['numbers_modified']:
                if numbers['old'] != numbers['new'] and numbers['old'] not in modified_numbers_list:
                    self.out.insert('end', '*** {} ===> {}\n'.format(numbers['old'], numbers['new']))
                    changes.write('*** {} ===> {}\n'.format(numbers['old'], numbers['new']))
                    modified_numbers_list.append(numbers['old'])

        changes.close()

    def compare_files(self):
        self.out.delete(1.0, 'end')
        result = self.controller.browse(
            self.controller.sheets_to_parse
        )
        if result:
            for r in result:
                self.out.insert('end', 'MATCH : {}\n'.format(r))
        else:
            self.out.insert('end', self.controller.sheets_to_parse)
