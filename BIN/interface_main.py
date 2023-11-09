#!/usr/bin/env python
#-*- coding:utf-8 -*-

import os, sys
from tkinter import *
from tkinter.font import Font
from tkinter.ttk import *
#Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel
from tkinter.messagebox import *
#Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')
import tkinter.filedialog as tkFileDialog
import tkinter.simpledialog as tkSimpleDialog  #askstring()
from PIL import Image, ImageTk
#from tkcalendar import Calendar, DateEntry
from language_dict import *


class Application_ui(Frame):
    #The class will create all widgets for UI.
    def __init__(self, LAN, master=None):
        self.LAN=LAN
        if self.LAN=='English':
            self.d=EngDict
        elif self.LAN=='日本語':
            self.d=JpnDict
        elif self.LAN=='ไทย':
            self.d=ThaDict
        Frame.__init__(self, master)
        self.master.title(self.d['C1'])
        self.master.geometry('919x598+141+20')
        self.createWidgets()

    def createWidgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.load = Image.open('AKB_Logo1.jpg')
        self.image = ImageTk.PhotoImage(self.load)
        self.AKB_Logo = Canvas(self.top, takefocus=1, highlightthickness=0)
        self.AKB_Logo.create_image(60, 50, image=self.image)
        self.AKB_Logo.place(relx=0.012, rely=0.82, relwidth=0.150, relheight=0.200)

        self.btnExitVar = StringVar(value=self.d['B4'])
        self.style.configure('TbtnExit.TButton', font=('Calibri',11))
        self.btnExit = Button(self.top, text=self.d['B4'], textvariable=self.btnExitVar, command=self.btnExit_Cmd, style='TbtnExit.TButton')
        self.btnExit.setText = lambda x: self.btnExitVar.set(x)
        self.btnExit.text = lambda : self.btnExitVar.get()
        self.btnExit.place(relx=0.766, rely=0.816, relwidth=0.175, relheight=0.055)

        self.btnGenerateVar = StringVar(value=self.d['B3'])
        self.style.configure('TbtnGenerate.TButton', font=('Calibri',11))
        self.btnGenerate = Button(self.top, text=self.d['B3'], textvariable=self.btnGenerateVar, command=self.btnGenerate_Cmd, style='TbtnGenerate.TButton')
        self.btnGenerate.setText = lambda x: self.btnGenerateVar.set(x)
        self.btnGenerate.text = lambda : self.btnGenerateVar.get()
        self.btnGenerate.place(relx=0.413, rely=0.816, relwidth=0.314, relheight=0.055)

        self.btnBatchVar = StringVar(value=self.d['B5'])
        self.style.configure('TbtnBatch.TButton', font=('Calibri', 11))
        self.btnBatch = Button(self.top, text=self.d['B5'], textvariable=self.btnBatchVar,
                                  command=self.btnBatch_Cmd, style='TbtnBatch.TButton')
        self.btnBatch.setText = lambda x: self.btnBatchVar.set(x)
        self.btnBatch.text = lambda: self.btnBatchVar.get()
        self.btnBatch.place(relx=0.153, rely=0.816, relwidth=0.214, relheight=0.055)

        self.ProgressBar1Var = StringVar(value='')
        self.ProgressBar1 = Progressbar(self.top, orient='horizontal', maximum=100, variable=self.ProgressBar1Var)
        self.ProgressBar1.place(relx=0.078, rely=0.722, relwidth=0.837, relheight=0.042)

        self.ProjectTVar = StringVar(value='')
        self.ProjectT = Entry(self.top, textvariable=self.ProjectTVar, font=('Calibri',12))
        self.ProjectT.setText = lambda x: self.ProjectTVar.set(x)
        self.ProjectT.text = lambda : self.ProjectTVar.get()
        self.ProjectT.place(relx=0.383, rely=0.575, relwidth=0.55, relheight=0.047)

        self.btnPersonVar = StringVar(value='...')
        self.style.configure('TbtnPerson.TButton', font=('Calibri',11))
        self.btnPerson = Button(self.top, text='...', textvariable=self.btnPersonVar, command=self.btnPerson_Cmd, style='TbtnPerson.TButton')
        self.btnPerson.setText = lambda x: self.btnPersonVar.set(x)
        self.btnPerson.text = lambda : self.btnPersonVar.get()
        self.btnPerson.place(relx=0.331, rely=0.575, relwidth=0.027, relheight=0.042)

        self.btnFolderVar = StringVar(value='...')
        self.style.configure('TbtnFolder.TButton', font=('Calibri',11))
        self.btnFolder = Button(self.top, text='...', textvariable=self.btnFolderVar, command=self.btnFolder_Cmd, style='TbtnFolder.TButton')
        self.btnFolder.setText = lambda x: self.btnFolderVar.set(x)
        self.btnFolder.text = lambda : self.btnFolderVar.get()
        self.btnFolder.place(relx=0.931, rely=0.375, relwidth=0.045, relheight=0.043)

        self.btnMasterVar = StringVar(value='...')
        self.style.configure('TbtnMaster.TButton', font=('Calibri',11))
        self.btnMaster = Button(self.top, text='...', textvariable=self.btnMasterVar, command=self.btnMaster_Cmd, style='TbtnMaster.TButton')
        self.btnMaster.setText = lambda x: self.btnMasterVar.set(x)
        self.btnMaster.text = lambda : self.btnMasterVar.get()
        self.btnMaster.place(relx=0.931, rely=0.268, relwidth=0.045, relheight=0.043)

        self.btnServerVar = StringVar(value='...')
        self.style.configure('TbtnServer.TButton', font=('Calibri',11))
        self.btnServer = Button(self.top, text='...', textvariable=self.btnServerVar, command=self.btnServer_Cmd, style='TbtnServer.TButton')
        self.btnServer.setText = lambda x: self.btnServerVar.set(x)
        self.btnServer.text = lambda : self.btnServerVar.get()
        self.btnServer.place(relx=0.931, rely=0.161, relwidth=0.045, relheight=0.043)

        self.btnCustomerVar = StringVar(value='...')
        self.style.configure('TbtnCustomer.TButton', font=('Calibri',11))
        self.btnCustomer = Button(self.top, text='...', textvariable=self.btnCustomerVar, command=self.btnCustomer_Cmd, style='TbtnCustomer.TButton')
        self.btnCustomer.setText = lambda x: self.btnCustomerVar.set(x)
        self.btnCustomer.text = lambda : self.btnCustomerVar.get()
        self.btnCustomer.place(relx=0.331, rely=0.294, relwidth=0.027, relheight=0.042)

        self.FolderTVar = StringVar(value='')
        self.FolderT = Entry(self.top, state='readonly', textvariable=self.FolderTVar, font=('Calibri',12))
        self.FolderT.setText = lambda x: self.FolderTVar.set(x)
        self.FolderT.text = lambda : self.FolderTVar.get()
        self.FolderT.place(relx=0.383, rely=0.375, relwidth=0.541, relheight=0.047)

        self.MasterTVar = StringVar(value='')
        self.MasterT = Entry(self.top, state='readonly', textvariable=self.MasterTVar, font=('Calibri',12))
        self.MasterT.setText = lambda x: self.MasterTVar.set(x)
        self.MasterT.text = lambda : self.MasterTVar.get()
        self.MasterT.place(relx=0.383, rely=0.268, relwidth=0.541, relheight=0.047)

        self.btnTestVar = StringVar(value=self.d['B2'])
        self.style.configure('TbtnTest.TButton', background='#FFFFFF', font=('Calibri', 11))
        self.btnTest = Button(self.top, text=self.d['B2'], textvariable=self.btnTestVar, command=self.btnTest_Cmd,
                              style='TbtnTest.TButton')
        self.btnTest.setText = lambda x: self.btnTestVar.set(x)
        self.btnTest.text = lambda: self.btnTestVar.get()
        self.btnTest.place(relx=0.522, rely=0.428, relwidth=0.271, relheight=0.057)

        self.ServerTVar = StringVar(value='')
        self.ServerT = Entry(self.top, state='readonly', background='white', textvariable=self.ServerTVar, font=('Calibri',12))
        self.ServerT.setText = lambda x: self.ServerTVar.set(x)
        self.ServerT.text = lambda : self.ServerTVar.get()
        self.ServerT.place(relx=0.383, rely=0.161, relwidth=0.541, relheight=0.047)

        self.SubmissionCList = ['In Progress','FIN','NO']
        self.SubmissionCVar = StringVar(value=self.d['C3'])
        self.SubmissionC = Combobox(self.top, text=self.d['C3'], textvariable=self.SubmissionCVar,state='readonly', values=self.SubmissionCList, font=('Calibri',12))
        self.SubmissionC.setText = lambda x: self.SubmissionCVar.set(x)
        self.SubmissionC.text = lambda : self.SubmissionCVar.get()
        self.SubmissionC.place(relx=0.044, rely=0.428, relwidth=0.28)
        self.SubmissionC.current(0)

        self.LanguageCList = ['English','日本語','ไทย']
        self.LanguageCVar = StringVar(value=self.LAN)
        self.LanguageC = Combobox(self.top, text=self.LAN, state='readonly', textvariable=self.LanguageCVar, values=self.LanguageCList, font=('Calibri',10))
        self.LanguageC.setText = lambda x: self.LanguageCVar.set(x)
        self.LanguageC.text = lambda : self.LanguageCVar.get()
        self.LanguageC.place(relx=0.879, rely=0.054, relwidth=0.097)

        self.CustomerCList = []
        self.CustomerCVar = StringVar(value=self.d['C2'])
        self.CustomerC = Combobox(self.top, text=self.d['C2'], state='readonly',textvariable=self.CustomerCVar, values=self.CustomerCList, font=('Calibri',12))
        self.CustomerC.setText = lambda x: self.CustomerCVar.set(x)
        self.CustomerC.text = lambda : self.CustomerCVar.get()
        self.CustomerC.place(relx=0.044, rely=0.294, relwidth=0.28)

        self.DateTVar = StringVar(value='19/11/2019')
        self.DateT = Entry(self.top, textvariable=self.DateTVar, font=('Calibri',12))
        self.DateT.setText = lambda x: self.DateTVar.set(x)
        self.DateT.text = lambda : self.DateTVar.get()
        self.DateT.place(relx=0.044, rely=0.161, relwidth=0.21, relheight=0.047)

        self.btnDateVar = StringVar(value=self.d['B1'])
        self.style.configure('TbtnDate.TButton', font=('Calibri',11))
        self.btnDate = Button(self.top, text=self.d['B1'], textvariable=self.btnDateVar, command=self.btnDate_Cmd, style='TbtnDate.TButton')
        self.btnDate.setText = lambda x: self.btnDateVar.set(x)
        self.btnDate.text = lambda : self.btnDateVar.get()
        self.btnDate.place(relx=0.261, rely=0.161, relwidth=0.097, relheight=0.042)

        self.PersonCList = []
        self.PersonCVar = StringVar(value=self.d['C4'])
        self.PersonC = Combobox(self.top, text=self.d['C4'], state='readonly', textvariable=self.PersonCVar, values=self.PersonCList, font=('Calibri',12))
        self.PersonC.setText = lambda x: self.PersonCVar.set(x)
        self.PersonC.text = lambda : self.PersonCVar.get()
        self.PersonC.place(relx=0.044, rely=0.575, relwidth=0.28)

        self.DateLVar = StringVar(value=self.d['L1'])
        self.style.configure('TDateL.TLabel', anchor='w', font=('Calibri',12))
        self.DateL = Label(self.top, text=self.d['L1'], textvariable=self.DateLVar, style='TDateL.TLabel')
        self.DateL.setText = lambda x: self.DateLVar.set(x)
        self.DateL.text = lambda : self.DateLVar.get()
        self.DateL.place(relx=0.044, rely=0.107, relwidth=0.193, relheight=0.042)

        self.Copyright_LabVar = StringVar(value=self.d['L10'])
        self.style.configure('TCopyright_Lab.TLabel', anchor='w', font=('Cambria',11,'italic'))
        self.Copyright_Lab = Label(self.top, text=self.d['L10'], textvariable=self.Copyright_LabVar, style='TCopyright_Lab.TLabel')
        self.Copyright_Lab.setText = lambda x: self.Copyright_LabVar.set(x)
        self.Copyright_Lab.text = lambda : self.Copyright_LabVar.get()
        self.Copyright_Lab.place(relx=0.242, rely=0.93, relwidth=0.68, relheight=0.055)

        self.MessageLVar = StringVar(value='')
        self.style.configure('TMessageL.TLabel', anchor='w',foreground='dark green', font=('Cambria',14,'bold'))
        self.MessageL = Label(self.top, text='', textvariable=self.MessageLVar, style='TMessageL.TLabel')
        self.MessageL.setText = lambda x: self.MessageLVar.set(x)
        self.MessageL.text = lambda : self.MessageLVar.get()
        self.MessageL.place(relx=0.044, rely=0.642, relwidth=0.832, relheight=0.055)

        self.ProjectLVar = StringVar(value=self.d['L8'])
        self.style.configure('TProjectL.TLabel', anchor='w', font=('Calibri',12))
        self.ProjectL = Label(self.top, text=self.d['L8'], textvariable=self.ProjectLVar, style='TProjectL.TLabel')
        self.ProjectL.setText = lambda x: self.ProjectLVar.set(x)
        self.ProjectL.text = lambda : self.ProjectLVar.get()
        self.ProjectL.place(relx=0.383, rely=0.508, relwidth=0.332, relheight=0.042)

        self.ServerLVar = StringVar(value=self.d['L5'])
        self.style.configure('TServerL.TLabel', anchor='w', font=('Calibri',12))
        self.ServerL = Label(self.top, text=self.d['L5'], textvariable=self.ServerLVar, style='TServerL.TLabel')
        self.ServerL.setText = lambda x: self.ServerLVar.set(x)
        self.ServerL.text = lambda : self.ServerLVar.get()
        self.ServerL.place(relx=0.383, rely=0.107, relwidth=0.193, relheight=0.042)

        self.PersonLVar = StringVar(value=self.d['L4'])
        self.style.configure('TPersonL.TLabel', anchor='w', font=('Calibri',12))
        self.PersonL = Label(self.top, text=self.d['L4'], textvariable=self.PersonLVar, style='TPersonL.TLabel')
        self.PersonL.setText = lambda x: self.PersonLVar.set(x)
        self.PersonL.text = lambda : self.PersonLVar.get()
        self.PersonL.place(relx=0.044, rely=0.508, relwidth=0.193, relheight=0.042)

        self.SubmissionLVar = StringVar(value=self.d['L3'])
        self.style.configure('TSubmissionL.TLabel', anchor='w', font=('Calibri',12))
        self.SubmissionL = Label(self.top, text=self.d['L3'], textvariable=self.SubmissionLVar, style='TSubmissionL.TLabel')
        self.SubmissionL.setText = lambda x: self.SubmissionLVar.set(x)
        self.SubmissionL.text = lambda : self.SubmissionLVar.get()
        self.SubmissionL.place(relx=0.044, rely=0.375, relwidth=0.193, relheight=0.042)

        self.CustomerLVar = StringVar(value=self.d['L2'])
        self.style.configure('TCustomerL.TLabel', anchor='w', font=('Calibri',12))
        self.CustomerL = Label(self.top, text=self.d['L2'], textvariable=self.CustomerLVar, style='TCustomerL.TLabel')
        self.CustomerL.setText = lambda x: self.CustomerLVar.set(x)
        self.CustomerL.text = lambda : self.CustomerLVar.get()
        self.CustomerL.place(relx=0.044, rely=0.241, relwidth=0.193, relheight=0.042)

        self.CheckVar1=IntVar()
        self.CheckVar2=StringVar(value=self.d['K1'])
        self.style.configure('TCheckBox.TLabel', anchor='w', font=('Calibri', 9))
        self.CheckBox=Checkbutton(self.top, text = self.d['K1'], textvariable=self.CheckVar2,\
                                  variable = self.CheckVar1, \
                 onvalue = 1, offvalue = 0)
        self.CheckBox.place(relx=0.244, rely=0.251, relwidth=0.123, relheight=0.042)

        self.FolderLVar = StringVar(value=self.d['L7'])
        self.style.configure('TFolderL.TLabel', anchor='w', font=('Calibri',12))
        self.FolderL = Label(self.top, text=self.d['L7'], textvariable=self.FolderLVar, style='TFolderL.TLabel')
        self.FolderL.setText = lambda x: self.FolderLVar.set(x)
        self.FolderL.text = lambda : self.FolderLVar.get()
        self.FolderL.place(relx=0.383, rely=0.321, relwidth=0.201, relheight=0.042)

        self.MasterLVar = StringVar(value=self.d['L6'])
        self.style.configure('TMasterL.TLabel', anchor='w', font=('Calibri',12))
        self.MasterL = Label(self.top, text=self.d['L6'], textvariable=self.MasterLVar, style='TMasterL.TLabel')
        self.MasterL.setText = lambda x: self.MasterLVar.set(x)
        self.MasterL.text = lambda : self.MasterLVar.get()
        self.MasterL.place(relx=0.383, rely=0.214, relwidth=0.193, relheight=0.042)

        self.LanguageLVar = StringVar(value=self.d['L9'])
        self.style.configure('TLanguageL.TLabel', anchor='w', font=('Calibri',10))
        self.LanguageL = Label(self.top, text=self.d['L9'], textvariable=self.LanguageLVar, style='TLanguageL.TLabel')
        self.LanguageL.setText = lambda x: self.LanguageLVar.set(x)
        self.LanguageL.text = lambda : self.LanguageLVar.get()
        self.LanguageL.place(relx=0.914, rely=0.013, relwidth=0.071, relheight=0.028)

        self.TitleLVar = StringVar(value=self.d['T1'])
        self.style.configure('TTitleL.TLabel', anchor='w', foreground='dark red', font=('Calisto MT',18,'bold'))
        self.TitleL = Label(self.top, text=self.d['T1'], textvariable=self.TitleLVar, style='TTitleL.TLabel')
        self.TitleL.setText = lambda x: self.TitleLVar.set(x)
        self.TitleL.text = lambda : self.TitleLVar.get()
        self.TitleL.place(relx=0.280, rely=0.027, relwidth=0.536, relheight=0.055)




