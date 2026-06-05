VERSION 5.00
Begin VB.Form FrmPodesavanje 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Podesavanja:"
   ClientHeight    =   4065
   ClientLeft      =   2280
   ClientTop       =   3450
   ClientWidth     =   3870
   Icon            =   "FrmPodesavanje.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   3870
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Zvono"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3615
      Begin BellPro.Slider sldDuzinaZvona 
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
      End
      Begin VB.Label lblDuzinaZvona 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Duzina Zvona"
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
         TabIndex        =   14
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   3615
      Begin VB.CheckBox chZvonoPrekoRazglasa 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Zvono preko razglasa"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox chVannastavne 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vannastavne aktivnosti"
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
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CheckBox chZastitaLozinkom 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Zastita lozinkom"
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
         TabIndex        =   3
         Top             =   840
         Width           =   3255
      End
      Begin VB.CheckBox chStartWithWin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pokreni program uz Windows"
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
         TabIndex        =   1
         Top             =   360
         Width           =   3255
      End
      Begin VB.CheckBox chObavestiZvono 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Obavesti kad zvoni"
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
         TabIndex        =   2
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox chStartMin 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Minimizuj pri startu"
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
         TabIndex        =   0
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   3615
      Begin VB.CheckBox chNeZvoni 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ne zvoni"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   2535
      End
      Begin VB.CheckBox chNeZvoniPraznici 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ne zvoni za praznike"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox chNeZvoniSubotom 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ne zvoni subotom "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chNeZvoniNedeljom 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ne zvoni nedeljom"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
      Begin VB.CheckBox chNeZvoniRaspust 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ne zvoni za raspust"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmPodesavanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ZastitaMenjanja As Boolean
Private Sub chZastitaLozinkom_Click()
    ZastitaMenjanja = True
End Sub
Private Sub Form_Load()
If Not DebugMode = True Then On Error Resume Next
    sldDuzinaZvona.value = GetProfile("config", "025", 8, getConfigPath)
    chStartWithWin.value = GetProfile("config", "026", chStartWithWin.value, getConfigPath)
    chObavestiZvono.value = GetProfile("config", "027", chObavestiZvono.value, getConfigPath)
    chNeZvoni.value = GetProfile("config", "028", chNeZvoni.value, getConfigPath)
    chStartMin.value = GetProfile("config", "029", chStartMin.value, getConfigPath)
    chNeZvoniSubotom.value = GetProfile("config", "030", chNeZvoniSubotom.value, getConfigPath)
    chNeZvoniNedeljom.value = GetProfile("config", "031", chNeZvoniNedeljom.value, getConfigPath)
    chNeZvoniRaspust.value = GetProfile("config", "032", chNeZvoniRaspust.value, getConfigPath)
    chNeZvoniPraznici.value = GetProfile("config", "034", chNeZvoniPraznici.value, getConfigPath)
    chZastitaLozinkom.value = GetProfile("config", "035", chZastitaLozinkom.value, getConfigPath)
    chVannastavne.value = GetProfile("config", "036", chVannastavne.value, getConfigPath)
    chZvonoPrekoRazglasa.value = GetProfile("config", "037", chZvonoPrekoRazglasa.value, getConfigPath)
    lblDuzinaZvona.Caption = "Duzina zvona: " & sldDuzinaZvona.value & " sec"
    ZastitaMenjanja = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "025", sldDuzinaZvona.value, getConfigPath
    WriteProfile "config", "026", chStartWithWin.value, getConfigPath
    WriteProfile "config", "027", chObavestiZvono.value, getConfigPath
    WriteProfile "config", "028", chNeZvoni.value, getConfigPath
    WriteProfile "config", "029", chStartMin.value, getConfigPath
    WriteProfile "config", "030", chNeZvoniSubotom.value, getConfigPath
    WriteProfile "config", "031", chNeZvoniNedeljom.value, getConfigPath
    WriteProfile "config", "032", chNeZvoniRaspust.value, getConfigPath
    WriteProfile "config", "034", chNeZvoniPraznici.value, getConfigPath
    WriteProfile "config", "035", chZastitaLozinkom.value, getConfigPath
    WriteProfile "config", "036", chVannastavne.value, getConfigPath
    WriteProfile "config", "037", chZvonoPrekoRazglasa.value, getConfigPath
    If chStartWithWin.value = 1 Then
        SetRegValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "BellPro", App.Path & "\" & App.EXEName & ".exe"
    Else
        DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "BellPro"
    End If
    If ZastitaMenjanja = True Then Call FrmMain.PostaviZastitu
    FrmMain.Show
    Unload Me
End Sub

Private Sub sldDuzinaZvona_Change(MyVal As Long, myMaxVal As Long)
If Not DebugMode = True Then On Error Resume Next
    lblDuzinaZvona.Caption = "Duzina zvona: " & sldDuzinaZvona.value & " sec"
End Sub

