VERSION 5.00
Begin VB.Form FrmInterfejs 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Interfejs:"
   ClientHeight    =   2925
   ClientLeft      =   -15
   ClientTop       =   255
   ClientWidth     =   2430
   Icon            =   "FrmInterfejs.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   2430
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbInterface 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1935
   End
   Begin VB.Timer Trm 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2640
      Top             =   1440
   End
   Begin VB.ComboBox CmbBrojPorta 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   1935
   End
   Begin BellPro.XPButton cmdTest 
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "TEST"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin BellPro.XPButton cmdRefresh 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Refresh "
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   120
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label lblBellInterfejs 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Bell interfejs: "
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
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label LblBrojPorta 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Broj porta: "
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
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "FrmInterfejs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmbBrojPorta_Change()
If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "023", CmbBrojPorta.Text, getConfigPath
End Sub

Private Sub CmbInterface_Change()
If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "024", CmbInterface.Text, getConfigPath
End Sub

Private Sub cmdRefresh_Click()
    GeneratePortList
End Sub
Private Sub CmdTest_Click()
If Not DebugMode = True Then On Error Resume Next
    Dim poruka As Integer
    poruka = MsgBox("Testiranjem interfejsa aktiviraæete relej na Interfejsu. Ukoliko je relej spojen sa sistememom zvona docice do oglašavanja zvona koje ce trajati 3 sec." & vbCrLf & "Da li zelite da odmah to uradite ?", vbQuestion & vbYesNo)
    If poruka = vbYes Then
        OpenPort
        HitTheRelay (True)
        Trm.Enabled = True
        cmdTest.Enabled = False
    End If
If poruka = vbNo Then Exit Sub
End Sub

Private Sub Form_Load()
On Error Resume Next
    GeneratePortList
    GenerateInterfaceList
    CmbBrojPorta.Text = GetProfile("config", "023", "COM1", getConfigPath)
    CmbInterface.Text = GetProfile("config", "024", "", getConfigPath)
End Sub
Private Sub GeneratePortList()
If Not DebugMode = True Then On Error Resume Next
    Dim i As Integer, x As Integer
    x = 0
    CmbBrojPorta.Clear
    For i = 1 To 17
        If IsCommExist(i) Then
            CmbBrojPorta.AddItem "COM" & i
            x = x + 1
        End If
    Next
    If x <> 0 Then CmbBrojPorta.ListIndex = 0
End Sub
Private Sub GenerateInterfaceList()
If Not DebugMode = True Then On Error Resume Next
    With CmbInterface
        .Clear
        .AddItem "Bell USB"
        .AddItem "Bell Ethernet"
        .AddItem "Bell Comm"
        .AddItem "Bez interfejsa"
        .ListIndex = 0
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "023", CmbBrojPorta.Text, getConfigPath
    WriteProfile "config", "024", CmbInterface.Text, getConfigPath
    MdlCommunication.ClosePort
    MdlCommunication.OpenPort
    FrmMain.Show
    Unload Me
End Sub

Private Sub trm_Timer()
If Not DebugMode = True Then On Error Resume Next
    HitTheRelay (False)
    cmdTest.Enabled = True
    MdlCommunication.ClosePort
    Trm.Enabled = False
End Sub
