#!/usr/bin/env python
#-*- coding:utf-8 -*-

import os, sys
from tkinter import *
from tkinter.font import Font
from tkinter.ttk import *

#Usage:showinfo/warning/error,askquestion/okcancel/yesno/retrycancel
from tkinter.messagebox import *
#Usage:f=tkFileDialog.askopenfilename(initialdir='E:/Python')
#import tkinter.filedialog as tkFileDialog
#import tkinter.simpledialog as tkSimpleDialog  #askstring()

class Application_ui(Frame):
    #The class will create all widgets for UI.
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('Form1')
        self.master.geometry('769x474+101+101')
        self.createWidgets()

    def createWidgets(self):
        self.top = self.winfo_toplevel()

        self.style = Style()

        self.List1Var = StringVar(value='List1')
        self.List1Font = Font(font=('SimSun',9))
        self.List1 = Listbox(self.top, listvariable=self.List1Var, font=self.List1Font)
        self.List1.place(relx=0.062, rely=0.152, relwidth=0.428, relheight=0.667)

        self.Command1Var = StringVar(value='Command1')
        self.style.configure('TCommand1.TButton', font=('SimSun',9))
        self.Command1 = Button(self.top, text='Command1', textvariable=self.Command1Var, command=self.Command1_Cmd, style='TCommand1.TButton')
        self.Command1.setText = lambda x: self.Command1Var.set(x)
        self.Command1.text = lambda : self.Command1Var.get()
        self.Command1.place(relx=0.135, rely=0.861, relwidth=0.241, relheight=0.07)

class Application_ui1(Frame):
    # The class will create all widgets for UI.
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.master.title('Form2')
        self.master.geometry('414x259')
        self.createWidgets()

    def createWidgets(self):
        self.top1 = self.winfo_toplevel()

        self.style = Style()

        self.Combo1List = ['Add items in designer or code!','2222','333','4444' ]
        self.Combo1Var = StringVar(value='Add items in designer or code!')
        self.Combo1 = Combobox(self.top1, text='Add items in designer or code!', textvariable=self.Combo1Var,
                                   values=self.Combo1List, font=('SimSun', 9))
        self.Combo1.setText = lambda x: self.Combo1Var.set(x)
        self.Combo1.text = lambda: self.Combo1Var.get()
        self.Combo1.place(relx=0.155, rely=0.185, relwidth=0.601)

        self.Command2Var = StringVar(value='Command1')
        self.style.configure('TCommand1.TButton', font=('SimSun',9))
        self.Command2 = Button(self.top1, text='Command1', textvariable=self.Command2Var, command=self.Command2_Cmd, style='TCommand1.TButton')
        self.Command2.setText = lambda x: self.Command2Var.set(x)
        self.Command2.text = lambda : self.Command2Var.get()
        self.Command2.place(relx=0.174, rely=0.618, relwidth=0.64, relheight=0.22)



class Application(Application_ui):
    #The class will implement callback function for events and your logical code.
    def __init__(self, master=None):
        Application_ui.__init__(self, master)

    def Command1_Cmd(self, event=None):
        def button_enabled():
            self.Command1['state'] = 'enabled'
            top1.destroy()

        class Application1(Application_ui1):

        # The class will implement callback function for events and your logical code.
            def __init__(self, master=None):
                Application_ui1.__init__(self, master)

            def Command2_Cmd(self, event=None):
                #self.Combo1.setText(5)
                showinfo(title='information',message=self.Combo1Var.get())


        self.Command1['state']='disabled'

        top1 = Toplevel()
        top1.protocol('WM_DELETE_WINDOW', button_enabled)
        Application1(top1).mainloop()


if __name__ == "__main__":
    top = Tk()
    Application(top).mainloop()

