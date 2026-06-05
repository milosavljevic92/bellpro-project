VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmSvakodnevni 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Upravljanje rasporedima:"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6120
   ControlBox      =   0   'False
   Icon            =   "FrmSvakodnevni.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdTestTime 
      Caption         =   "TEST TIME"
      Height          =   735
      Left            =   360
      TabIndex        =   14
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin BellPro.Slider sldLen 
      Height          =   285
      Left            =   2400
      TabIndex        =   7
      Top             =   6360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   503
   End
   Begin VB.ComboBox CmbRaspored 
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
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   240
      Width           =   2295
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
      Left            =   120
      MaxLength       =   55
      TabIndex        =   0
      Top             =   5880
      Width           =   4815
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
      Left            =   5040
      MaxLength       =   5
      TabIndex        =   1
      Top             =   5880
      Width           =   975
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   7858
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
   Begin BellPro.XPButton cmdObrisiVreme 
      Height          =   615
      Left            =   2760
      TabIndex        =   9
      Top             =   6960
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
      ForeColor       =   255
      ForeHover       =   0
   End
   Begin BellPro.XPButton cmdSacuvaj 
      Height          =   615
      Left            =   1440
      TabIndex        =   10
      Top             =   6960
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
      Top             =   6960
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
      Left            =   4800
      TabIndex        =   12
      Top             =   6960
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
   Begin BellPro.XPButton cmdNovi 
      Height          =   375
      Left            =   2640
      TabIndex        =   13
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Dodaj raspored"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin BellPro.XPButton cmdObrisiRaspored 
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Obrisi raspored"
      ForeColor       =   192
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
      TabIndex        =   6
      Top             =   6360
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   6000
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Shape Shape1 
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblNazivZvona 
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
      Left            =   120
      TabIndex        =   4
      Top             =   5520
      Width           =   4815
   End
   Begin VB.Label lblVreme 
      Alignment       =   2  'Center
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
      Left            =   5040
      TabIndex        =   3
      Top             =   5520
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   6000
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "FrmSvakodnevni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private raspored As ADODB.Recordset
Private tabela As ADODB.Recordset
Private ComboRaspored As ADODB.Recordset
Dim InsertNew As Boolean
Private Sub CmbRaspored_Change()
If Not DebugMode = True Then On Error Resume Next
    Call PostaviGrid(CmbRaspored.Text)
    Me.TxtNaziv.Text = tabela.Fields("Naziv").value
    Me.TxtVreme.Text = tabela.Fields("Vreme").value
    sldLen.value = ReturnNonAlpha(tabela.Fields("DuzinaZvona").value)
    lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
End Sub
Private Sub CmbRaspored_Click()
If Not DebugMode = True Then On Error Resume Next
Call PostaviGrid(CmbRaspored.Text)
    If tabela.RecordCount <> 0 Then
        Me.TxtNaziv.Text = tabela.Fields("Naziv").value
        Me.TxtVreme.Text = tabela.Fields("Vreme").value
        sldLen.value = ReturnNonAlpha(tabela.Fields("DuzinaZvona").value)
        lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    End If
End Sub
Private Sub CmdDodaj_Click()
If Not DebugMode = True Then On Error Resume Next
    cmdObrisiVreme.Enabled = False
    cmdDodaj.Enabled = False
    cmdNazad.Enabled = False
    cmdSacuvaj.Enabled = True
    TxtNaziv.Text = ""
    TxtVreme.Text = ""
    sldLen.value = GetProfile("config", "025", 8, getConfigPath)
    lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    tabela.AddNew
    InsertNew = True
End Sub
Private Sub CmdNazad_Click()
If Not DebugMode = True Then On Error Resume Next
    FrmMain.PostaviCombo
    FrmMain.PostaviGridMain (GetProfile("config", "021", "", getConfigPath))
    FrmMain.Show
    Unload Me
End Sub

Private Sub cmdNovi_Click()
    Me.Hide
    FrmNoviRaspored.Show
End Sub

Private Sub cmdObrisiRaspored_Click()
If Not DebugMode = True Then On Error Resume Next
    Dim odgovor
    If CmbRaspored.Text = "Svakodnevni" Then
        MsgBox "Svakodnevni raspored je sistemski raspored i ne moze biti obrisan!", vbInformation
        Exit Sub
    End If
    If GetProfile("config", "021", "", getConfigPath) = CmbRaspored.Text Then
        odgovor = MsgBox("Raspored koji zelite da obrisete je trenutno u upotrebi!" & vbNewLine & "Da li sigurno zelite obrisati raspored " & Me.CmbRaspored.Text & " ?", vbQuestion & vbYesNo)
        If odgovor = vbYes Then Call ObrisiRaspored(CmbRaspored.Text)
        If odgovor = vbNo Then Exit Sub
    Else
        odgovor = MsgBox("Da li sigurno zelite obrisati raspored " & Me.CmbRaspored.Text & " ?", vbYesNo)
        If odgovor = vbYes Then Call ObrisiRaspored(CmbRaspored.Text)
        If odgovor = vbNo Then Exit Sub
    End If
End Sub

Private Sub cmdObrisiVreme_Click()
    If Not DebugMode = True Then On Error Resume Next
    If tabela.RecordCount = 0 Then Exit Sub
    Dim obrisi As Integer
    obrisi = MsgBox("Dali ste sigurni da zelite obrisati vreme?", vbQuestion & vbYesNo)
    If obrisi = vbYes Then tabela.Delete
    If obrisi = vbNo Then Exit Sub
    If tabela.RecordCount = 0 Then
        Me.TxtNaziv.Text = ""
        Me.TxtVreme.Text = ""
        sldLen.value = GetProfile("config", "025", 8, getConfigPath)
        lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
        cmdObrisiVreme.Enabled = False
    End If
End Sub

Private Sub CmdSacuvaj_Click()
If Not DebugMode = True Then On Error Resume Next
    If ProveriFormatVremena(TxtVreme.Text) = False Then
        MsgBox "Vreme formatirano nepravilno!", vbCritical
        Exit Sub
    End If
    If tabela.RecordCount = 0 Then Exit Sub
    cmdObrisiVreme.Enabled = True
    cmdDodaj.Enabled = True
    cmdNazad.Enabled = True
    tabela.Fields("Naziv").value = Me.TxtNaziv.Text
    tabela.Fields("Dan").value = "pon - ned"
    tabela.Fields("Vreme").value = Me.TxtVreme.Text & ":00"
    tabela.Fields("Raspored").value = CmbRaspored.Text
    tabela.Fields("DuzinaZvona").value = sldLen.value & " sec"
    tabela.Update
    InsertNew = False
    MsgBox "Uspesno sacuvano!", vbInformation
End Sub
Private Sub genTestRecord()
    tabela.MoveFirst

    Dim x As Integer
    Dim i As Integer
    For i = 0 To 23
        For x = 1 To 59
            tabela.AddNew
            tabela.Fields("Naziv").value = "Test Time"
            tabela.Fields("Vreme").value = format(i, "00") + ":" + format(x, "00") + ":00"
            tabela.Fields("Dan").value = "pon - ned"
            tabela.Fields("Raspored").value = CmbRaspored.Text
            tabela.Fields("DuzinaZvona").value = "15 sec"
            tabela.Update
        Next x
    Next i
End Sub

Private Sub cmdTestTime_Click()
genTestRecord
End Sub

Private Sub Form_Load()
If Not DebugMode = True Then On Error Resume Next
If DebugMode = True Then cmdTestTime.Visible = True
    Call PostaviComboPomocni
    CmbRaspored.Text = GetProfile("config", "021", "", getConfigPath)
    Call PostaviGrid(GetProfile("config", "021", "", getConfigPath))
    If tabela.RecordCount <> 0 Then
        Me.TxtNaziv.Text = tabela.Fields("Naziv").value
        Me.TxtVreme.Text = tabela.Fields("Vreme").value
        Me.sldLen.value = ReturnNonAlpha(tabela.Fields("DuzinaZvona").value)
        lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    Else
        cmdSacuvaj.Enabled = False
        cmdObrisiVreme.Enabled = False
    End If
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
On Error Resume Next
    If tabela.RecordCount <> 0 And InsertNew = False Then
        Me.TxtNaziv.Text = tabela.Fields("Naziv").value
        Me.TxtVreme.Text = tabela.Fields("Vreme").value
        sldLen.value = ReturnNonAlpha(tabela.Fields("DuzinaZvona").value)
        lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    End If
End Sub

Public Function KreirajNoviRaspored(Naziv As String) As Boolean
If Not DebugMode = True Then On Error Resume Next
Set raspored = New ADODB.Recordset
  raspored.Open "SELECT * FROM NaziviRasporeda WHERE Naziv='" & Naziv & "'", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
  If raspored.RecordCount = 0 Then
        raspored.AddNew
        raspored.Fields("Naziv").value = Naziv
        raspored.Update
        KreirajNoviRaspored = True
        Call PostaviComboPomocni
        PostaviGrid (Naziv)
        CmbRaspored.Text = Naziv
  Else
        KreirajNoviRaspored = False
  End If
  raspored.Close
End Function
Private Sub PostaviComboPomocni()
If Not DebugMode = True Then On Error Resume Next
    Dim x As Integer
    CmbRaspored.Clear
    Set ComboRaspored = New ADODB.Recordset
    ComboRaspored.Open "SELECT DISTINCT Naziv FROM NaziviRasporeda", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    For x = 1 To ComboRaspored.RecordCount
       CmbRaspored.AddItem ComboRaspored.Fields("Naziv").value
       ComboRaspored.MoveNext
    Next x
    ComboRaspored.Close
End Sub
Private Sub PostaviGrid(NazivRasporeda As String)
    If Not DebugMode = True Then On Error Resume Next
    Set tabela = New ADODB.Recordset
    TxtNaziv.Text = ""
    TxtVreme.Text = ""
    sldLen.value = GetProfile("config", "025", 8, getConfigPath)
    lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
    tabela.Open "SELECT * FROM raspored WHERE Raspored='" & NazivRasporeda & "'", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    If tabela.RecordCount = 0 Then
        tabela.Close
        NazivRasporeda = "PraznoPolje"
        tabela.Open "SELECT * FROM raspored WHERE Raspored='" & NazivRasporeda & "'", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    End If
    Set Grid.DataSource = tabela
    Grid.Refresh
    Grid.MarqueeStyle = dbgHighlightRow
    Grid.Columns.Remove ("ID")
    Grid.Columns.Remove ("Raspored")
    Grid.Columns.Remove ("Dan")
    Grid.Columns("Vreme").Width = 1250
    Grid.Columns("Naziv").Width = 2950
    Grid.Columns(2).Caption = "Duzina Zvona"
    Grid.Columns(2).Width = 1250
    Grid.Columns(2).Alignment = dbgCenter
End Sub
Private Sub ObrisiRaspored(NazivRasporeda As String)
If Not DebugMode = True Then On Error Resume Next
    Set raspored = New ADODB.Recordset
    Dim x As Integer
    raspored.Open "SELECT * FROM Raspored WHERE Raspored='" & NazivRasporeda & "'", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    If Not raspored.RecordCount = 0 Then
          For x = 1 To raspored.RecordCount
            raspored.Delete
            raspored.MoveNext
          Next x
    End If
    raspored.Close
    raspored.Open "SELECT * FROM NaziviRasporeda WHERE Naziv='" & NazivRasporeda & "'", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    raspored.Delete
    raspored.Close
    raspored.Open "SELECT * FROM NaziviRasporeda", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    If Not raspored.RecordCount = 0 Then
        raspored.MoveFirst
        Call PostaviComboPomocni
        Call PostaviGrid(raspored.Fields("Naziv").value)
        CmbRaspored.Text = raspored.Fields("Naziv").value
    Else
        Call PostaviComboPomocni
        CmbRaspored.Text = " "
        
    End If
    raspored.Close
End Sub


Private Sub sldLen_Change(MyVal As Long, myMaxVal As Long)
  lblLen.Caption = "Duzina zvona: " & sldLen.value & " sec"
End Sub
