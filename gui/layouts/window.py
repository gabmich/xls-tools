# coding: utf-8

import tkinter as tk
from tkinter.messagebox import *
from tkinter.filedialog import *
from parsers import XlsDuplicateParser as xd
from editors import PhonenumberParser as pp

class Launcher:
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.add_button({'text':'XLS Duplicate Parser', 'width':25, 'command':lambda: self.open_app(xd)})
        self.add_button({'text':'XLS Phonenumbers Editor', 'width':25, 'command':lambda: self.open_app(pp)})
        self.add_button({'text':'QUIT', 'width':25, 'command':lambda: self.close_windows()})
        self.frame.pack()

    def add_button(self, kwargs):
        """ Add a button to the main app window """
        self.button = tk.Button(self.frame, **kwargs)
        self.button.pack()

    def open_app(self, app):
        """ Open a sub app window """
        self.newWindow = tk.Toplevel(self.master)
        self.app = app(self.newWindow)

    def close_windows(self):
        """ Close an app window """
        self.master.destroy()