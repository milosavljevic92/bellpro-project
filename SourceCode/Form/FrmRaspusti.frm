VERSION 5.00
Begin VB.Form FrmRaspusti 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Zadavanje raspusta:"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4710
   Icon            =   "FrmRaspusti.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text8 
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
      Left            =   3000
      TabIndex        =   15
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text7 
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
      Left            =   1200
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text5 
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
      Left            =   1200
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text6 
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
      Left            =   3000
      TabIndex        =   12
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text3 
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
      Left            =   1200
      TabIndex        =   11
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text4 
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
      Left            =   3000
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox Text2 
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
      Left            =   3000
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox Text1 
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
      Left            =   1200
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "do"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
   Begin VB.Shape Shape2 
      Height          =   615
      Left            =   120
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "do"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Raspust 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Raspust 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Shape Shape3 
      Height          =   615
      Left            =   120
      Top             =   1560
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "do"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Raspust 3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1800
      Width           =   975
   End
   Begin VB.Shape Shape4 
      Height          =   615
      Left            =   120
      Top             =   2280
      Width           =   4455
   End
   Begin VB.Label Label10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "do"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   255
   End
   Begin VB.Label Label12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Raspust 4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
End
Attribute VB_Name = "FrmRaspusti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    If Not DebugMode = True Then On Error Resume Next
    
   Text1.Text = GetProfile("config", "013", "", getConfigPath)
    Text2.Text = GetProfile("config", "014", "", getConfigPath)
    Text3.Text = GetProfile("config", "015", "", getConfigPath)
    Text4.Text = GetProfile("config", "016", "", getConfigPath)
    Text5.Text = GetProfile("config", "017", "", getConfigPath)
    Text6.Text = GetProfile("config", "018", "", getConfigPath)
    Text7.Text = GetProfile("config", "019", "", getConfigPath)
    Text8.Text = GetProfile("config", "020", "", getConfigPath)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "013", Text1.Text, getConfigPath
    WriteProfile "config", "014", Text2.Text, getConfigPath
    WriteProfile "config", "015", Text3.Text, getConfigPath
    WriteProfile "config", "016", Text4.Text, getConfigPath
    WriteProfile "config", "017", Text5.Text, getConfigPath
    WriteProfile "config", "018", Text6.Text, getConfigPath
    WriteProfile "config", "019", Text7.Text, getConfigPath
    WriteProfile "config", "020", Text8.Text, getConfigPath
    FrmMain.Show
    Unload Me
End Sub
