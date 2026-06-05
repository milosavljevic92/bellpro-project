VERSION 5.00
Begin VB.Form FrmRucnoZ 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Rucno upravljanje zvonom:"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2910
   Icon            =   "FrmRucnoZ.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   2910
   StartUpPosition =   2  'CenterScreen
   Begin BellPro.XPButton cmdRucnoZvoni 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2655
      _extentx        =   4683
      _extenty        =   1508
      font            =   "FrmRucnoZ.frx":802D
      caption         =   "Upali Zvono"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin VB.ComboBox CmbNacinZvona 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   120
      X2              =   2760
      Y1              =   720
      Y2              =   720
   End
End
Attribute VB_Name = "FrmRucnoZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdRucnoZvoni_Click()
'If Not DebugMode = True Then On Error Resume Next
    Select Case cmdRucnoZvoni.Caption
    Case "Upali Zvono"
        cmdRucnoZvoni.Caption = "Ugasi Zvono"
        CmbNacinZvona.Enabled = False
        If CmbNacinZvona.Text = "Zvono" Then HitTheRelay (True)
        If CmbNacinZvona.Text = "Razglas" Then PlayMp3Sound
        If CmbNacinZvona.Text = "Zvono + Razglas" Then
            HitTheRelay (True)
            PlayMp3Sound
        End If
    Case "Ugasi Zvono"
        cmdRucnoZvoni.Caption = "Upali Zvono"
         CmbNacinZvona.Enabled = True
        If CmbNacinZvona.Text = "Zvono" Then HitTheRelay (False)
        If CmbNacinZvona.Text = "Razglas" Then StopMp3Sound
        If CmbNacinZvona.Text = "Zvono + Razglas" Then
            HitTheRelay (False)
            StopMp3Sound
        End If
    End Select
End Sub

Private Sub Form_Load()
CmbNacinZvona.AddItem "Zvono"
CmbNacinZvona.AddItem "Razglas"
CmbNacinZvona.AddItem "Zvono + Razglas"
CmbNacinZvona.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
HitTheRelay (False)
StopMp3Sound
Unload Me
End Sub
