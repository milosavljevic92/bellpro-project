VERSION 5.00
Begin VB.Form FrmRegistracija 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " BellPro | Registracija programa"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   Icon            =   "FrmRegistracija.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin BellPro.XPButton cmdDemo 
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "DEMO"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin BellPro.XPButton cmdCopy 
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Copy"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.TextBox TxtImeskole 
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
      Left            =   360
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Width           =   3975
   End
   Begin VB.TextBox TxtParametar 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   600
      Width           =   2535
   End
   Begin VB.TextBox TxtLicenca 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   3975
   End
   Begin VB.PictureBox Picture1 
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   0
      TabIndex        =   4
      Top             =   0
      Width           =   0
   End
   Begin BellPro.XPButton cmdRegistruj 
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Registruj"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Label lblNazivUstanove 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Naziv ustanove:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label LblParametar 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Parametar:"
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
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label LblLicenca 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Licenca:"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00400000&
      BorderWidth     =   2
      Height          =   3015
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "FrmRegistracija"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
If Not DebugMode = True Then On Error Resume Next
    Dim licenca As String
    TxtLicenca.Text = Kriptuj(GetSetting("Tecomatic", "BellPro", "002", ""), False)
    TxtParametar.Text = Trim$(VolumeSerialNumber(Left(App.Path, 3)))
    setAboutApp False
    If GenLicencu(TxtParametar.Text) = TxtLicenca.Text Then
        FrmSplash.Show
        Unload Me
    Else
        TxtLicenca.Text = ""
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    Unload Me
End Sub
Private Sub CmdCopy_Click()
On Error Resume Next
Clipboard.Clear
Clipboard.SetText TxtParametar.Text
End Sub
Private Sub cmdDemo_Click()
    setDemoMode
    SaveSetting "Tecomatic", "BellPro", "001", Kriptuj("1111", True)
    FrmSplash.Show
    Unload Me
End Sub
Private Sub CmdRegistruj_Click()
If Not DebugMode = True Then On Error Resume Next
    If GenLicencu(TxtParametar.Text) = TxtLicenca.Text Then
        SaveSetting "Tecomatic", "BellPro", "001", Kriptuj("1111", True)
        SaveSetting "Tecomatic", "BellPro", "002", Kriptuj(TxtLicenca.Text, True)
        SaveSetting "Tecomatic", "BellPro", "003", TxtImeskole.Text
        SaveSetting "Tecomatic", "BellPro", "004", format(Now, "dd.mm.yyyy")
        setFullMode
        MsgBox "Program uspesno registrovan, hvala sto se odabrali BellPro!", vbInformation
        FrmSplash.Show
        Unload Me
    Else
        MsgBox "Licenca koju ste uneli nije ispravna, mozda ste je pogresno uneli istu." _
        + vbNewLine + "Kontakt: tecomatic@gmail.com", vbCritical
        TxtLicenca.Text = ""
        setDemoMode
        Exit Sub
    End If
End Sub
Private Function GenLicencu(parametar As String) As String
If Not DebugMode = True Then On Error Resume Next
    Dim output As String, licenca As String
    output = EncodeADFGVX(TxtParametar.Text, "BellPro1511", "tecomatic")
    If output <> "" Then
        licenca = MakeGroups(output, True, Val(10))
        GenLicencu = licenca
    Else
        MsgBox "Nastala je greska pri registraciji programa, program æe se automatski zatvoriti. Pokusajte ponovo da ga registrujete."
        End
    End If
End Function



