# coding: utf-8

import tkinter as tk
from gui.window import Window
from gui.parsers import XlsDuplicateParserGUI as xd
from gui.editors import PhonenumberParserGUI as pp

root = tk.Tk()


class Launcher(Window):
    def __init__(self, window):
        Window.__init__(self, root)
        self.add_button({
            'text': 'XLS Duplicate Parser',
            'width': 25,
            'command': lambda: self.open_app(xd)
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


if __name__ == '__main__':
    app = Launcher(Window(root))
    root.mainloop()
