# coding: utf-8

import tkinter as tk
from tkinter import Grid

from gui import PnParserGUI


root = tk.Tk()

Grid.rowconfigure(root, 0, weight=1)
Grid.columnconfigure(root, 0, weight=1)
PnParserGUI(root)

root.update()
root.minsize(root.winfo_width(), root.winfo_height())

root.mainloop()