#!/usr/bin/env python
#-*- coding:utf-8 -*-

import os, sys
from tkinter import *
from tkinter.font import Font
from tkinter.ttk import *
from language_dict import *
#Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel
from tkinter.messagebox import *
#Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')
import tkinter.filedialog as tkFileDialog
import tkinter.simpledialog as tkSimpleDialog  #askstring()

class Customer_ui(Frame):
    #The class will create all widgets for UI.
    def __init__(self,LAN, master=None):
        self.LAN=LAN
        if self.LAN == 'English':
            self.d = EngDict_C
        elif self.LAN == '日本語':
            self.d = JpnDict_C
        elif self.LAN == 'ไทย':
            self.d = ThaDict_C
        Frame.__init__(self, master)
        self.master.title(self.d['T1'])
        self.master.geometry('482x460+300+80')
        self.createWidgets()

    def createWidgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.ClistVar = StringVar(value='')
        self.ClistFont = Font(font=('Calibri',12))
        self.Clist = Listbox(self.top, listvariable=self.ClistVar, font=self.ClistFont)
        self.Clist.place(relx=0.066, rely=0.157, relwidth=0.55, relheight=0.793)
#--------------Scroll Bar
        self.yscrollbar = Scrollbar(self.Clist, command=self.Clist.yview)
        self.yscrollbar.pack(side=RIGHT, fill=Y)
        self.Clist.config(yscrollcommand=self.yscrollbar.set)


        self.InputBoxVar = StringVar(value='')
        self.InputBox = Entry(self.top, textvariable=self.InputBoxVar, font=('Calibri',12))
        self.InputBox.setText = lambda x: self.InputBoxVar.set(x)
        self.InputBox.text = lambda : self.InputBoxVar.get()
        self.InputBox.place(relx=0.066, rely=0.052, relwidth=0.55, relheight=0.072)

        self.btnAddCVar = StringVar(value=self.d['B1'])
        self.style.configure('TbtnAddC.TButton', font=('SimSun',10))
        self.btnAddC = Button(self.top, text=self.d['B1'], textvariable=self.btnAddCVar, command=self.btnAddC_Cmd, style='TbtnAddC.TButton')
        self.btnAddC.setText = lambda x: self.btnAddCVar.set(x)
        self.btnAddC.text = lambda : self.btnAddCVar.get()
        self.btnAddC.place(relx=0.664, rely=0.052, relwidth=0.284, relheight=0.072)

        self.btnDeleteCVar = StringVar(value=self.d['B2'])
        self.style.configure('TbtnDeleteC.TButton', font=('SimSun',10))
        self.btnDeleteC = Button(self.top, text=self.d['B2'], textvariable=self.btnDeleteCVar, command=self.btnDeleteC_Cmd, style='TbtnDeleteC.TButton')
        self.btnDeleteC.setText = lambda x: self.btnDeleteCVar.set(x)
        self.btnDeleteC.text = lambda : self.btnDeleteCVar.get()
        self.btnDeleteC.place(relx=0.664, rely=0.452, relwidth=0.284, relheight=0.072)

        self.btnResumeVar = StringVar(value=self.d['B3'])
        self.style.configure('TbtnResume.TButton', font=('SimSun',10))
        self.btnResume = Button(self.top, text=self.d['B3'], textvariable=self.btnResumeVar, command=self.btnResume_Cmd, style='TbtnResume.TButton')
        self.btnResume.setText = lambda x: self.btnResumeVar.set(x)
        self.btnResume.text = lambda : self.btnResumeVar.get()
        self.btnResume.place(relx=0.664, rely=0.87, relwidth=0.284, relheight=0.072)

        self.btnOKVar = StringVar(value=self.d['B4'])
        self.style.configure('TbtnOK.TButton', font=('SimSun',10))
        self.btnOK = Button(self.top, text=self.d['B4'], textvariable=self.btnOKVar, command=self.btnOK_Cmd, style='TbtnOK.TButton')
        self.btnOK.setText = lambda x: self.btnOKVar.set(x)
        self.btnOK.text = lambda : self.btnOKVar.get()
        self.btnOK.place(relx=0.664, rely=0.77, relwidth=0.284, relheight=0.072)

