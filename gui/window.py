# coding: utf-8

import tkinter as tk


class Window():
    def __init__(self, master):
        self.master = master
        self.frame = tk.Frame(self.master)

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
