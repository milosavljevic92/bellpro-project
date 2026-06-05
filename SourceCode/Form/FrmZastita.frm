VERSION 5.00
Begin VB.Form FrmZastita 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Promena PIN-a"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2805
   Icon            =   "FrmZastita.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   2805
   StartUpPosition =   1  'CenterOwner
   Begin BellPro.XPButton cmdSacuvaj 
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   2535
      _extentx        =   4471
      _extenty        =   1085
      font            =   "FrmZastita.frx":802D
      caption         =   "Sacuvaj"
      forecolor       =   -2147483642
      forehover       =   0
   End
   Begin VB.TextBox TxtPonovoNova 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   4
      PasswordChar    =   "•"
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox TxtNova 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   4
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox TxtTrenutna 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   4
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   120
      X2              =   2640
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblPonoviNoviPin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Potvrdi novi pin: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblNoviPin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Novi pin:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label lblTrenutniPin 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Trenutni pin [4 broja]:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmZastita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdSacuvaj_Click()
If Not DebugMode = True Then On Error Resume Next
    Dim lozinka As String
    lozinka = Kriptuj(GetSetting("Tecomatic", "BellPro", "001", ""), False)
    If Not TxtTrenutna.Text = lozinka Then
        MsgBox "Trenutni PIN koji ste uneli je neispravan!", vbCritical
        Call ClearTextBoxes
        Exit Sub
    Else
        If TxtNova.Text = "" And TxtPonovoNova.Text = "" Then
            MsgBox "Niste potvrdili novi PIN!", vbCritical
            Call ClearTextBoxes
            Exit Sub
        Else
                If Not TxtNova.Text = TxtPonovoNova.Text Then
                    MsgBox "PIN nije isti u poljima za potvrdu novog!", vbCritical
                    Call ClearTextBoxes
                    Exit Sub
                Else
                    SaveSetting "Tecomatic", "BellPro", "001", Kriptuj(TxtNova.Text, True)
                    MsgBox "Novi PIN postavljen!", vbInformation
                    FrmMain.Show
                    Unload Me
                    Exit Sub
                End If
        End If
    End If
End Sub
Private Sub ClearTextBoxes()
Me.TxtNova.Text = ""
Me.TxtPonovoNova.Text = ""
Me.TxtTrenutna.Text = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    FrmMain.Show
    Unload Me
End Sub

Private Sub TxtNova_KeyPress(KeyAscii As Integer)
If Not DebugMode = True Then On Error Resume Next
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtTrenutna_KeyPress(KeyAscii As Integer)
If Not DebugMode = True Then On Error Resume Next
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub TxtPonovoNova_KeyPress(KeyAscii As Integer)
If Not DebugMode = True Then On Error Resume Next
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If
End Sub
