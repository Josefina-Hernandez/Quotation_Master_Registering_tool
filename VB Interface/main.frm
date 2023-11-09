VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form Main_Form 
   Caption         =   "AKT Quotation Manager(Ver.1.1)"
   ClientHeight    =   8970
   ClientLeft      =   1635
   ClientTop       =   1980
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   13785
   Begin VB.PictureBox AKB_Logo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1545
      Left            =   120
      Picture         =   "main.frx":0000
      ScaleHeight     =   103
      ScaleMode       =   0  'User
      ScaleWidth      =   118
      TabIndex        =   29
      Top             =   7200
      Width           =   1770
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Quit System"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10560
      TabIndex        =   28
      Top             =   7320
      Width           =   2415
   End
   Begin VB.CommandButton btnGenerate 
      Caption         =   "Generate Project"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   27
      Top             =   7320
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   1080
      TabIndex        =   25
      Top             =   6480
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox ProjectT 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      TabIndex        =   24
      Text            =   "CLAAS_Translation of Brochure LEXION 6000 LRC (JP)"
      Top             =   5160
      Width           =   7575
   End
   Begin VB.CommandButton btnPerson 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton btnFolder 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12840
      TabIndex        =   20
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton btnMaster 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12840
      TabIndex        =   19
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton btnServer 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12840
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton btnCustomer 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   17
      Top             =   2640
      Width           =   375
   End
   Begin VB.TextBox FolderT 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      TabIndex        =   16
      Text            =   "\\192.168.11.10\PM_secretB\2020\Quotation Request"
      Top             =   3360
      Width           =   7455
   End
   Begin VB.TextBox MasterT 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      TabIndex        =   14
      Text            =   "//192.168.11.10/Sales_secretB/TEST2020.xls"
      Top             =   2400
      Width           =   7455
   End
   Begin VB.CommandButton btnTest 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Connection Test"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7200
      MaskColor       =   &H00C0FFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      UseMaskColor    =   -1  'True
      Width           =   3735
   End
   Begin VB.TextBox ServerT 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5280
      TabIndex        =   11
      Text            =   "192.168.11.10"
      Top             =   1440
      Width           =   7455
   End
   Begin VB.ComboBox SubmissionC 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   8
      Text            =   "In Progress"
      Top             =   3840
      Width           =   4335
   End
   Begin VB.ComboBox LanguageC 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   12120
      TabIndex        =   6
      Text            =   "English"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox CustomerC 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   4
      Text            =   "TAKIGEN"
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox DateT 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   600
      TabIndex        =   2
      Text            =   "19/11/2019"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton btnDate 
      Caption         =   "Date Select"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox PersonC 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      TabIndex        =   10
      Text            =   "Itoh"
      Top             =   5160
      Width           =   3855
   End
   Begin VB.Label DateL 
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   31
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Copyright_Lab 
      Caption         =   "All Copyrights belong to AKAGANE Business Support Co., Ltd          Developed in 2019  Ver.1.1"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   30
      Top             =   8160
      Width           =   9375
   End
   Begin VB.Label MessageL 
      Caption         =   "Message Area"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   26
      Top             =   5760
      Width           =   4575
   End
   Begin VB.Label ProjectL 
      Caption         =   "Project Name (Free Entry)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   23
      Top             =   4560
      Width           =   4575
   End
   Begin VB.Label ServerL 
      Caption         =   "Server IP Address"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   21
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label PersonL 
      Caption         =   "Person In Charge"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label SubmissionL 
      Caption         =   "Submission"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Label CustomerL 
      Caption         =   "Customer Select"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label FolderL 
      Caption         =   "PM Folder Path"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   15
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label MasterL 
      Caption         =   "Quotation Master Path"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label LanguageL 
      Caption         =   "LANGUAGE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   12600
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label TitleL 
      Caption         =   "AKB Quotation Manager (Ver.1.1)"
      BeginProperty Font 
         Name            =   "Calisto MT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "Main_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
