VERSION 5.00
Begin VB.Form Customer_List 
   Caption         =   "Customer List"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7230
   LinkTopic       =   "Form2"
   ScaleHeight     =   6900
   ScaleWidth      =   7230
   StartUpPosition =   3  '‚xå˚„ûè»
   Begin VB.CommandButton btnResume 
      Caption         =   "Resume"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   4
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton btnDeleteC 
      Caption         =   "Delete Customer"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3120
      Width           =   2055
   End
   Begin VB.CommandButton btnAddC 
      Caption         =   "Add Customer"
      BeginProperty Font 
         Name            =   "SimSun"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.TextBox InputBox 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.ListBox Clist 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      ItemData        =   "AKT_Quotation_Manager.frx":0000
      Left            =   480
      List            =   "AKT_Quotation_Manager.frx":0007
      TabIndex        =   0
      Top             =   1080
      Width           =   3975
   End
End
Attribute VB_Name = "Customer_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub
