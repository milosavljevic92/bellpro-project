VERSION 5.00
Begin VB.Form FrmOtkljucavanje 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Unesite PIN: "
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4215
   Icon            =   "FrmOtkljucavanje.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TrmEnter 
      Interval        =   50
      Left            =   3360
      Top             =   720
   End
   Begin VB.TextBox TxtLozinka 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   4
      PasswordChar    =   "•"
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
   Begin BellPro.XPButton cmdPrijavise 
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1335
      _extentx        =   2355
      _extenty        =   873
      font            =   "FrmOtkljucavanje.frx":802D
      caption         =   "Prijavi se"
      forecolor       =   -2147483642
      forehover       =   0
   End
End
Attribute VB_Name = "FrmOtkljucavanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Dim textval As String
Dim numval As String
Private Sub CmdOdustani_Click()

End Sub
Private Sub CmdPrijavise_Click()
If Not DebugMode = True Then On Error Resume Next
Dim lozinka As String
TrmEnter.Enabled = False
lozinka = Kriptuj(GetSetting("Tecomatic", "BellPro", "001", ""), False)

If TxtLozinka.Text = lozinka Then
    FrmMain.lockApp True
    TxtLozinka.Text = ""
    FrmMain.Show
    Unload Me
Else
    MsgBox "Lozinka nije ispravana!", vbCritical
    TxtLozinka.Text = ""
    TxtLozinka.TabIndex = 0
End If
TrmEnter.Enabled = True
End Sub
Private Sub Form_Load()
If Not DebugMode = True Then On Error Resume Next
TxtLozinka.TabIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
FrmMain.Show
Unload Me
End Sub

Private Sub TxtLozinka_KeyPress(KeyAscii As Integer)
If Not DebugMode = True Then On Error Resume Next
    If KeyAscii = 13 Then
        CmdPrijavise_Click
    End If
    If Not IsNumeric(Chr(KeyAscii)) And Not KeyAscii = 8 Then
        KeyAscii = 0
    End If

End Sub
