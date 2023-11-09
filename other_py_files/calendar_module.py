import tkinter as tk
from tkinter import ttk

from tkcalendar import Calendar

class Date_Select():
    def __init__(self,master):
        self.master=master
        self.top=tk.Toplevel(master=master)
        self.struct_calender()
        self.returnVal=''

    def print_sel(self,event):
        self.returnVal=self.cal.selection_get()
        global G
        G=self.returnVal
        print(G)

    def struct_calender(self):
        self.cal = Calendar(self.top,
                       font="Arial 14", selectmode='day',
                       cursor="hand1")  # , year=2018, month=2, day=5)
        self.cal.bind("<<CalendarSelected>>", self.print_sel)
        self.cal.pack(fill="both", expand=True)
        ttk.Button(self.top, text="ok", command=(self.top.destroy)).pack()




