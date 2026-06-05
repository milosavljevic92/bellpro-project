VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmVanNastavne 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Vannastavne aktivnosti:"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7815
   ControlBox      =   0   'False
   Icon            =   "FrmVanNastavne.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin BellPro.XPButton cmdObrisi 
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Obrisi"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin BellPro.Slider sldLen 
      Height          =   285
      Left            =   2520
      TabIndex        =   8
      Top             =   6240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   503
   End
   Begin VB.TextBox TxtVreme 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      MaxLength       =   5
      TabIndex        =   2
      Top             =   5760
      Width           =   975
   End
   Begin VB.TextBox TxtNaziv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      MaxLength       =   55
      TabIndex        =   1
      Top             =   5760
      Width           =   4815
   End
   Begin VB.ComboBox CmbDan 
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
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   14737632
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2074
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2074
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin BellPro.XPButton cmdSacuvaj 
      Height          =   615
      Left            =   1440
      TabIndex        =   10
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Sacuvaj"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin BellPro.XPButton cmdDodaj 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Dodaj"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin BellPro.XPButton cmdNazad 
      Height          =   615
      Left            =   6480
      TabIndex        =   12
      Top             =   6840
      Width           =   1215
      _ExtentX        =   2143
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
      Caption         =   "Nazad"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Label lblLen 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Duzina zvona: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
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
      Top             =   6240
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7680
      Y1              =   5280
      Y2              =   5280
   End
   Begin VB.Label lblVreme 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vreme:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   6720
      TabIndex        =   5
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label lblOpisZvona 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Opis zvona:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   1800
      TabIndex        =   4
      Top             =   5400
      Width           =   4815
   End
   Begin VB.Label lblDan 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Dan:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   1215
   End
End
Attribute VB_Name = "FrmVanNastavne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private raspored As ADODB.Recordset
Private Sub CmdDodaj_Click()
If Not DebugMode = True Then On Error Resume Next
    raspored.AddNew
    Me.cmdObrisi.Enabled = False
    Me.cmdDodaj.Enabled = False
    Me.cmdNazad.Enabled = False
    Me.cmdSacuvaj.Enabled = True
    Me.TxtNaziv.Text = ""
    Me.TxtVreme.Text = ""
    sldLen.value = GetProfile("config", "025", 8, getConfigPath)
    lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
End Sub
Private Sub CmdNazad_Click()
If Not DebugMode = True Then On Error Resume Next
    Unload Me
    FrmMain.Show
End Sub
Private Sub cmdObrisi_Click()
If Not DebugMode = True Then On Error Resume Next
    If raspored.RecordCount = 0 Then Exit Sub
    Dim poruka As Integer
    poruka = MsgBox("Dali ste sigurni da zelite obrisati vreme?", vbQuestion & vbYesNo)
    If poruka = vbYes Then raspored.Delete
    If poruka = vbNo Then Exit Sub
        If raspored.RecordCount = 0 Then
            Me.TxtNaziv.Text = ""
            Me.TxtVreme.Text = ""
            Me.CmbDan.ListIndex = 0
            sldLen.value = GetProfile("config", "025", 8, getConfigPath)
            lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
            cmdObrisi.Enabled = False
            cmdSacuvaj.Enabled = False
        End If
End Sub

Private Sub CmdSacuvaj_Click()
If Not DebugMode = True Then On Error Resume Next
    If ProveriFormatVremena(TxtVreme.Text) = False Then
        MsgBox "Vreme formatirano nepravilno!", vbCritical
        Exit Sub
    End If
    Me.cmdObrisi.Enabled = True
    Me.cmdDodaj.Enabled = True
    Me.cmdNazad.Enabled = True
    raspored.Fields("Naziv").value = Me.TxtNaziv.Text
    raspored.Fields("Dan").value = Me.CmbDan.Text
    raspored.Fields("Vreme").value = Me.TxtVreme.Text & ":00"
    raspored.Fields("Raspored").value = "VanNastavni"
    raspored.Fields("DuzinaZvona").value = sldLen.value & " sec"
    raspored.Update
End Sub
Private Sub Form_Load()
If Not DebugMode = True Then On Error Resume Next
    PopuniCombo
    PostaviGrid ("SELECT * FROM raspored WHERE Raspored='VanNastavni'")
End Sub
Private Sub PopuniCombo()
    With CmbDan
        .AddItem " "
        .AddItem "Ponedeljak"
        .AddItem "Utorak"
        .AddItem "Sreda"
        .AddItem "Cetvrtak"
        .AddItem "Petak"
        .AddItem "Subota"
        .AddItem "Nedelja"
        .ListIndex = 0
    End With
End Sub
Private Sub PostaviGrid(query As String)
    Set raspored = New ADODB.Recordset
    raspored.Open query, FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    With Grid
        Set .DataSource = raspored
        .Refresh
        .MarqueeStyle = dbgHighlightRow
        .Columns.Remove ("ID")
        .Columns.Remove ("Raspored")
        .Columns("Vreme").Width = 1000
        .Columns("dan").Width = 1150
        .Columns("Naziv").Width = 3300
        .Columns(3).Caption = "Duzina Zvona"
        .Columns(3).Width = 1250
        .Columns(3).Alignment = dbgCenter
    End With
    sldLen.value = GetProfile("config", "025", 8, getConfigPath)
    lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    If raspored.RecordCount <> 0 Then
        Me.TxtNaziv.Text = raspored.Fields("Naziv").value
        Me.CmbDan.Text = raspored.Fields("Dan").value
        Me.TxtVreme.Text = raspored.Fields("Vreme").value
        sldLen.value = ReturnNonAlpha(raspored.Fields("DuzinaZvona").value)
        lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    Else
        cmdSacuvaj.Enabled = False
        Me.cmdObrisi.Enabled = False
    End If
End Sub
Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If raspored.RecordCount <> 0 Then
        Me.TxtNaziv.Text = raspored.Fields("Naziv").value
        Me.CmbDan.Text = raspored.Fields("Dan").value
        Me.TxtVreme.Text = raspored.Fields("Vreme").value
        sldLen.value = ReturnNonAlpha(raspored.Fields("DuzinaZvona").value)
        lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    End If
End Sub
Private Sub sldLen_Change(MyVal As Long, myMaxVal As Long)
    lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
End Sub

