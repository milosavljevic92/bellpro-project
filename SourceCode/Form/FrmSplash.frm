VERSION 5.00
Begin VB.Form FrmSplash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3285
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5910
   ControlBox      =   0   'False
   Icon            =   "FrmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TrmSplash 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   7560
      Top             =   720
   End
   Begin VB.Label lblDatumInstalacije 
      BackStyle       =   0  'Transparent
      Caption         =   "Registrovano od:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label lblClose 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BellPro - School version"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5655
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1215
      Left            =   120
      Picture         =   "FrmSplash.frx":802D
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   5655
   End
   Begin VB.Label LblInt 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5655
   End
   Begin VB.Label LblNaziv 
      BackStyle       =   0  'Transparent
      Caption         =   "Licenca"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Support: office@tecomatic.net"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5895
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Year: 2009 - 2020"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5895
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xx As Single
Dim yy As Single
Private Sub Form_Load()
If Not DebugMode = False Then On Error Resume Next
    Me.Caption = ""
    LblInt.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    LblNaziv.Caption = "User: " & GetSetting("Tecomatic", "BellPro", "003", "")
    lblDatumInstalacije.Caption = "Date: " & GetSetting("Tecomatic", "BellPro", "004", "")

    If Not AboutApp = True Then
        
        TrmSplash.Enabled = True
        Me.lblClose.Visible = False
    Else
        Me.lblClose.Visible = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    FrmMain.Show
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub TrmSplash_Timer()
If Not DebugMode = True Then On Error Resume Next
    TrmSplash.Enabled = False
    Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    xx = x
    yy = Y
End Sub

