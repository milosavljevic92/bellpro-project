VERSION 5.00
Begin VB.Form FrmNoviRaspored 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Novi raspored:"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6075
   Icon            =   "FrmNoviRaspored.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtNaziv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      MaxLength       =   20
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin BellPro.XPButton cmdNapravi 
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   480
      Width           =   1455
      _extentx        =   2566
      _extenty        =   873
      font            =   "FrmNoviRaspored.frx":802D
      caption         =   "Napravi"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin VB.Label lblNazivRasporeda 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Naziv rasporeda:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "FrmNoviRaspored"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdNapravi_Click()
If Not DebugMode = True Then On Error Resume Next
    Dim rez As Boolean
    rez = FrmSvakodnevni.KreirajNoviRaspored(TxtNaziv.Text)
    If rez = False Then
        MsgBox "Ime rasporeda vec postoji u bazi!", vbCritical
        Exit Sub
    Else
        FrmSvakodnevni.Show
        Unload Me
    End If
End Sub

Private Sub cmdRefresh_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    FrmSvakodnevni.Show
    Unload Me
End Sub
