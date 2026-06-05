VERSION 5.00
Begin VB.Form FrmPraznik 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Definisanje praznika:"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6645
   Icon            =   "FmrPraznik.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   6645
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text20 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   11
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text19 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "9# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4320
      TabIndex        =   23
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "10# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   4320
      TabIndex        =   22
      Top             =   840
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "11# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   4320
      TabIndex        =   21
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "12# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   4320
      TabIndex        =   20
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "5# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2280
      TabIndex        =   19
      Top             =   240
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "6# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   2280
      TabIndex        =   18
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "7# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   17
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "8# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2280
      TabIndex        =   16
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "4# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "3# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "2# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   840
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "1# datum:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FrmPraznik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    If Not DebugMode = True Then On Error Resume Next
    Text9.Text = GetProfile("config", "001", Text9.Text, getConfigPath)
    Text10.Text = GetProfile("config", "002", Text10.Text, getConfigPath)
    Text11.Text = GetProfile("config", "003", Text11.Text, getConfigPath)
    Text12.Text = GetProfile("config", "004", Text12.Text, getConfigPath)
    Text13.Text = GetProfile("config", "005", Text13.Text, getConfigPath)
    Text14.Text = GetProfile("config", "006", Text14.Text, getConfigPath)
    Text15.Text = GetProfile("config", "007", Text15.Text, getConfigPath)
    Text16.Text = GetProfile("config", "008", Text16.Text, getConfigPath)
    Text17.Text = GetProfile("config", "009", Text17.Text, getConfigPath)
    Text18.Text = GetProfile("config", "010", Text18.Text, getConfigPath)
    Text19.Text = GetProfile("config", "011", Text19.Text, getConfigPath)
    Text20.Text = GetProfile("config", "012", Text20.Text, getConfigPath)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "001", Text9.Text, getConfigPath
    WriteProfile "config", "002", Text10.Text, getConfigPath
    WriteProfile "config", "003", Text11.Text, getConfigPath
    WriteProfile "config", "004", Text12.Text, getConfigPath
    WriteProfile "config", "005", Text13.Text, getConfigPath
    WriteProfile "config", "006", Text14.Text, getConfigPath
    WriteProfile "config", "007", Text15.Text, getConfigPath
    WriteProfile "config", "008", Text16.Text, getConfigPath
    WriteProfile "config", "009", Text17.Text, getConfigPath
    WriteProfile "config", "010", Text18.Text, getConfigPath
    WriteProfile "config", "011", Text19.Text, getConfigPath
    WriteProfile "config", "012", Text20.Text, getConfigPath
    FrmMain.Show
    Unload Me
End Sub
