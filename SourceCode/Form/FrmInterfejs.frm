VERSION 5.00
Begin VB.Form FrmInterfejs 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "BellPro Interfejs:"
   ClientHeight    =   3510
   ClientLeft      =   -15
   ClientTop       =   255
   ClientWidth     =   2970
   Icon            =   "FrmInterfejs.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2970
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCommPort 
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Text            =   "COM1"
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CheckBox chDtrCtsEnabled 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Detekcija Interfejsa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CheckBox chInvertRTS 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Invertuj rad releja"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   2775
   End
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
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin VB.Timer Trm 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   0
      Top             =   3000
   End
   Begin BellPro.XPButton cmdTest 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   2760
      Width           =   2775
      _ExtentX        =   4895
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
      ForeColor       =   192
      ForeHover       =   0
   End
   Begin BellPro.XPButton cmdRefresh 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      Enabled         =   0   'False
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
      Width           =   2775
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
      TabIndex        =   2
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
      TabIndex        =   0
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
Private Sub Form_Load()
    ApplyLanguage Me
On Error Resume Next
    GenerateInterfaceList
    txtCommPort.Text = GetProfile("config", "023", "COM1", getConfigPath)
    CmbInterface.Text = GetProfile("config", "024", "", getConfigPath)
    chDtrCtsEnabled.value = GetProfile("config", "040", chDtrCtsEnabled.value, getConfigPath)
    chInvertRTS.value = GetProfile("config", "041", chInvertRTS.value, getConfigPath)
End Sub
Private Sub CmbInterface_Change()
If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "024", CmbInterface.Text, getConfigPath
End Sub
Private Sub GenerateInterfaceList()
If Not DebugMode = True Then On Error Resume Next
    With CmbInterface
        .Clear
        .AddItem t("FrmInterfejs", "ItemBellCommDirect")
        .AddItem t("FrmInterfejs", "ItemBellProRelay")
        .AddItem t("FrmInterfejs", "ItemBezInterfejsa")
        .ListIndex = 0
    End With
End Sub

Private Sub CmdTest_Click()
If Not DebugMode = True Then On Error Resume Next
    Dim poruka As Integer
    poruka = MsgBox(t("FrmInterfejs", "MsgTestInterfejsa"), vbQuestion & vbYesNo)
    If poruka = vbYes Then
        ClosePort
        OpenPort
        HitTheRelay (True)
        Trm.Enabled = True
        cmdTest.Enabled = False
    End If
    If poruka = vbNo Then Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "023", txtCommPort.Text, getConfigPath
    WriteProfile "config", "024", CmbInterface.Text, getConfigPath
    WriteProfile "config", "040", chDtrCtsEnabled.value, getConfigPath
    WriteProfile "config", "041", chInvertRTS.value, getConfigPath
    FrmMain.Show
    Unload Me
End Sub

Private Sub trm_Timer()
If Not DebugMode = True Then On Error Resume Next
    Trm.Enabled = False
    HitTheRelay (False)
    cmdTest.Enabled = True
End Sub
