# coding: utf-8

import tkinter as tk
from tkinter import Grid

from gui import XlsPhonenumbersParserGUI


root = tk.Tk()

Grid.rowconfigure(root, 0, weight=1)
Grid.columnconfigure(root, 0, weight=1)
XlsPhonenumbersParserGUI(root)

root.mainloop()