# coding: utf-8

import tkinter as tk
from app.XlsDuplicateParser.gui import XlsDuplicateParserGUI as xdgui
from app.XlsDuplicateParser.controller import XlsDuplicateParser as xd
from app.editors import PhonenumberParserGUI as pp

root = tk.Tk()


class Launcher():
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)
        self.add_button({
            'text': 'XLS Duplicate Parser',
            'width': 25,
            'command': lambda: self.open_app(xdgui)
        })
        self.add_button({
            'text': 'XLS Phonenumbers Editor',
            'width': 25,
            'command': lambda: self.open_app(pp)
        })
        self.add_button({
            'text': 'QUIT',
            'width': 25,
            'command': lambda: self.close_windows()
        })
        self.frame.pack()

    def add_button(self, kwargs):
        """ Add a button to the main app window """
        self.button = tk.Button(self.frame, **kwargs)
        self.button.pack()

    def open_app(self, app):
        """ Open a sub app window """
        self.new_window = tk.Toplevel(self.master)
        self.app = app(self.new_window)

    def close_windows(self):
        """ Close an app window """
        self.master.destroy()


if __name__ == '__main__':
    app = Launcher(root)
    root.mainloop()
