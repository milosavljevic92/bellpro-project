VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "-"
   ClientHeight    =   5685
   ClientLeft      =   1680
   ClientTop       =   1980
   ClientWidth     =   5490
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   379
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   366
   StartUpPosition =   2  'CenterScreen
   Begin BellPro.CntrlSerial CntrlSerial 
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin VB.Timer trmTime 
      Interval        =   10
      Left            =   6480
      Top             =   840
   End
   Begin VB.Timer trmStatus 
      Interval        =   100
      Left            =   6960
      Top             =   840
   End
   Begin VB.ComboBox CmbRaspored 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Rasporedi definisani u bazi podataka"
      Top             =   120
      Width           =   2295
   End
   Begin VB.Timer TrmProvera 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   240
   End
   Begin VB.Timer TrmZvoni 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   6960
      Top             =   240
   End
   Begin VB.PictureBox StatusBar1 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   30
      Left            =   0
      ScaleHeight     =   30
      ScaleWidth      =   5490
      TabIndex        =   0
      Top             =   5655
      Width           =   5490
   End
   Begin MSDataGridLib.DataGrid Grid 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      Appearance      =   0
      BackColor       =   14737632
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
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
   Begin BellPro.XPButton cmdPrimeni 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Primeni"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Vreme : 02:00:00 | Datum: 30.11.2019 | Danas je: Nedelja | Tecomatic.net"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5400
      Width           =   5535
   End
   Begin VB.Image ImgDissconn 
      Height          =   375
      Left            =   4320
      Picture         =   "FrmMain.frx":802D
      Stretch         =   -1  'True
      ToolTipText     =   "Port je otvoren, ali interfejs nije pronadjen"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image ImgConnected 
      Height          =   375
      Left            =   4320
      Picture         =   "FrmMain.frx":B04B
      Stretch         =   -1  'True
      ToolTipText     =   "Interfejs BellFace je konektovan"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image ImgZelena 
      Height          =   375
      Left            =   4920
      Picture         =   "FrmMain.frx":EAF9
      Stretch         =   -1  'True
      ToolTipText     =   "Danas zvoni "
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Image ImgCrvena 
      Height          =   375
      Left            =   4920
      Picture         =   "FrmMain.frx":F353
      Stretch         =   -1  'True
      ToolTipText     =   "Danas ne zvoni"
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuRaspored 
      Caption         =   "&Raspored"
      Begin VB.Menu mnusvakodnevni 
         Caption         =   "&Izmena "
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuvannastavne 
         Caption         =   "&Vannastavne aktivnosti"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnucrta2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuUvezi 
         Caption         =   "&Uvezi"
      End
      Begin VB.Menu MnuIzvezi 
         Caption         =   "&Izvezi"
      End
   End
   Begin VB.Menu mnupodesavanja 
      Caption         =   "&Podesavanja"
      Begin VB.Menu mnupodprograma 
         Caption         =   "&Programa "
      End
      Begin VB.Menu mnuraspusti 
         Caption         =   "&Raspusti"
      End
      Begin VB.Menu mnupraznici 
         Caption         =   "&Praznici "
      End
      Begin VB.Menu mnuzastitap 
         Caption         =   "&Zastita"
      End
      Begin VB.Menu mnucrta 
         Caption         =   "-"
      End
      Begin VB.Menu mnuinterfejs 
         Caption         =   "&Interfejs"
      End
   End
   Begin VB.Menu mnuprogramu 
      Caption         =   "&Program"
      Begin VB.Menu mnuruucno 
         Caption         =   "&Rucno zvono"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuzastitaz 
         Caption         =   "&Zakljucaj"
      End
      Begin VB.Menu mnuupdate 
         Caption         =   "&Update"
         Visible         =   0   'False
      End
      Begin VB.Menu mnucrta1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About..."
      End
      Begin VB.Menu mnutray 
         Caption         =   "&Zatvori prozor"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuRegistracija 
      Caption         =   "&Registracija"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuizlaz 
      Caption         =   "&Izlaz"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Konekcija As ADODB.Connection
Private raspored As ADODB.Recordset, tabela As ADODB.Recordset, uvoz As ADODB.Recordset
Public vreme As String, datum As String
Private DanasZvoni As Boolean, DanasJeDan As String
Private Sub Form_Load()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    Dim boolTemp As Boolean
    Dim demoString As String
    Set Konekcija = New ADODB.Connection
    boolTemp = checkDoesDBexist(App.Path & "\base.sqlite")
    If boolTemp = False Then
        MsgBox "Baza podataka ne postoji!" & vbNewLine & "Pozovite tehnicku podrsku kako bih resili ovaj problem.", vbCritical
        End
    End If
    Konekcija.CursorLocation = adUseClient
    Konekcija.Open "Driver={SQLite3 ODBC Driver};Database=" & App.Path & "\base.sqlite;" & ";"
    
    With IconData
        .cbSize = Len(IconData)
        .hIcon = Me.Icon
        .hwnd = Me.hwnd
        .szTip = "Bell Pro " & Chr(0)
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uID = vbNull
    End With
    Call PostaviCombo
    Call PostaviGridMain(GetProfile("config", "021", "", getConfigPath))
    Call PostaviZastitu
    demoString = ""
    If getDemoStatus = True Then
        demoString = "DEMO"
        mnuRegistracija.Visible = True
    Else
        Call OpenPort
    End If
   
    Me.Caption = "BellPro - School Edition - V" & App.Major & "." & App.Minor & "." & App.Revision & " " & demoString
    If GetProfile("config", "029", "1", getConfigPath) = 1 Then
        Me.WindowState = 1
    Else
        Me.WindowState = 0
    End If
    TrmProvera.Enabled = True
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Not DebugMode = True Then On Error Resume Next
    Dim Msg As Long
    Msg = x
    If Msg = WM_LBUTTONDBLCLK Then Call mnuShow_Click
End Sub
Private Sub Form_Resize()
If Not DebugMode = True Then On Error Resume Next
    If Me.WindowState = 1 Then
        Call Shell_NotifyIcon(NIM_ADD, IconData)
        App.TaskVisible = False
        Me.Hide
    End If
End Sub
Private Sub CmdPrimeni_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    WriteProfile "config", "021", CmbRaspored.Text, getConfigPath
    Call PostaviGridMain(GetProfile("config", "021", "", getConfigPath))
    MsgBox "Raspored je primenjen", vbInformation
End Sub
Private Sub ImgCrvena_Click()
    FrmPodesavanje.Show
    FrmMain.Hide
End Sub
Private Sub ImgDissconn_Click()
    ClosePort
    FrmInterfejs.Show
    FrmMain.Hide
End Sub
Private Sub ImgZelena_Click()
    FrmPodesavanje.Show
    FrmMain.Hide
End Sub
Private Sub mnuAbout_Click()
    setAboutApp True
    FrmSplash.Show
End Sub
Private Sub mnuizlaz_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    Dim izlaz As String
    izlaz = MsgBox("Da li ste sigurni da zelite zatvoriti program?" + vbNewLine + "Ukoliko ga zatvorite zvono se nece oglasavati.", vbQuestion & vbYesNo)
    If izlaz = vbYes Then
        ClosePort
        Shell_NotifyIcon NIM_DELETE, IconData
        End
    End If
    If izlaz = vbNo Then Exit Sub
End Sub
Private Sub CmdInterfejs_Click()
If Not DebugMode = True Then On Error Resume Next
    FrmInterfejs.Show
    Me.Hide
End Sub
Private Sub mnuInterfejs_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    ClosePort
    FrmInterfejs.Show
    FrmMain.Hide
End Sub
Private Sub MnuIzvezi_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    FrmExport.Show
    FrmMain.Hide
End Sub
Private Sub mnupodprograma_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    FrmPodesavanje.Show
    FrmMain.Hide
End Sub
Private Sub mnupraznici_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    FrmPraznik.Show
    FrmMain.Hide
End Sub
Private Sub mnuraspusti_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    FrmRaspusti.Show
    FrmMain.Hide
End Sub

Private Sub mnuRegistracija_Click()
FrmRegistracija.Show
Unload Me
End Sub
Private Sub mnuruucno_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    FrmRucnoZ.Show
End Sub
Private Sub mnusvakodnevni_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    FrmSvakodnevni.Show
    FrmMain.Hide
End Sub
Private Sub mnutray_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    Me.WindowState = 1
End Sub
Private Sub mnuvannastavne_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    FrmVanNastavne.Show
    FrmMain.Hide
End Sub
Private Sub mnuzastitap_Click()
If Not DebugMode = True Then On Error Resume Next
    FrmZastita.Show
    FrmMain.Hide
End Sub
Private Sub mnuzastitaz_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    Select Case mnuzastitaz.Caption
    Case "Otkljucaj"
        FrmOtkljucavanje.Show
    Case "Zakljucaj"
        mnuzastitaz.Caption = "Otkljucaj"
        lockApp False
    End Select
End Sub
Private Sub mnuShow_Click()
If Not DebugMode = True Then On Error Resume Next
    FrmMain.WindowState = 0
    FrmMain.Show
    Shell_NotifyIcon NIM_DELETE, IconData
    App.TaskVisible = True
End Sub
Private Sub MnuUvezi_Click()
If Not DebugMode = True Then If Not DebugMode = True Then On Error Resume Next
    Dim OpenFile As OPENFILENAME
    Dim lReturn As Long
    Dim sFilter As String
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = FrmMain.hwnd
    OpenFile.hInstance = App.hInstance
    sFilter = "eXtensible Markup Language (*.xml)" & Chr(0) & "*.XML" & Chr(0)
    OpenFile.lpstrFilter = sFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrInitialDir = "C:\"
    OpenFile.lpstrTitle = "Ucitaj XML file..."
    OpenFile.flags = 0
    lReturn = GetOpenFileName(OpenFile)
    If lReturn = 0 Then
        Exit Sub
    Else
        CitajXml (OpenFile.lpstrFile)
    End If
End Sub

Private Sub TrmProvera_Timer()
If Not DebugMode = True Then On Error Resume Next
    Set raspored = New ADODB.Recordset
    Call setInterfacePic
    If ProveriSve = False Then
        DanasZvoni = True
        ImgZelena.Visible = True
        ImgCrvena.Visible = False
    Else
        DanasZvoni = False
        ImgZelena.Visible = False
        ImgCrvena.Visible = True
    End If
    
    raspored.Open "SELECT * FROM Raspored WHERE Raspored='" & GetProfile("config", "021", "", getConfigPath) & "' AND Vreme='" & vreme & "'", Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    If raspored.RecordCount <> 0 Then
        TrmProvera.Enabled = False
        Call UpaliZvono(raspored.Fields("Vreme").value, raspored.Fields("Naziv").value, ReturnNonAlpha(raspored.Fields("DuzinaZvona").value))
    End If
    
    raspored.Close
    If GetProfile("config", "036", "1", getConfigPath) = 1 Then
        raspored.Open "SELECT * FROM Raspored WHERE Raspored='VanNastavni' AND Vreme='" & vreme & "' AND Dan='" & DanasJeDan & "'", Konekcija, adOpenStatic, adLockOptimistic, adCmdText
        If raspored.RecordCount <> 0 Then
            TrmProvera.Enabled = False
            TrmZvoni.Interval = ReturnNonAlpha(raspored.Fields("DuzinaZvona").value) * 1000
            HitTheRelay (True)
            TrmZvoni.Enabled = True
        End If
        raspored.Close
    End If
End Sub
Private Sub UpaliZvono(VremeZvona As String, NazivZvona As String, DuzinaZvona As String)
If Not DebugMode = True Then On Error Resume Next
If DanasZvoni = True Then
    If GetProfile("config", "027", "1", getConfigPath) = 1 Then
        PrikaziPoruku "Zvono za vreme: " & VremeZvona & vbNewLine & "Opis: " & NazivZvona, DuzinaZvona
    End If
    HitTheRelay (True)
    If GetProfile("config", "037", "0", getConfigPath) = 1 Then PlayMp3Sound
    TrmZvoni.Interval = DuzinaZvona * 1000
    TrmZvoni.Enabled = True
Else
    TrmProvera.Enabled = True
End If

End Sub

Private Sub trmTime_Timer()
    datum = format(Now, "dd.MM.yyyy")
    vreme = format(Now, "hh:mm:ss")
End Sub
Private Sub TrmZvoni_Timer()
If Not DebugMode = True Then On Error Resume Next
    TrmZvoni.Enabled = False
    HitTheRelay (False)
    StopMp3Sound
    TrmProvera.Enabled = True
End Sub
Private Sub trmStatus_Timer()
    Call PostaviStatusBar
End Sub
Private Sub PostaviStatusBar()
If Not DebugMode = True Then On Error Resume Next
    If Weekday(Date) = 1 Then DanasJeDan = "Nedelja"
    If Weekday(Date) = 2 Then DanasJeDan = "Ponedeljak"
    If Weekday(Date) = 3 Then DanasJeDan = "Utorak"
    If Weekday(Date) = 4 Then DanasJeDan = "Sreda"
    If Weekday(Date) = 5 Then DanasJeDan = "Cetvrtak"
    If Weekday(Date) = 6 Then DanasJeDan = "Petak"
    If Weekday(Date) = 7 Then DanasJeDan = "Subota"
    If Weekday(Date) = 1 Then DanasJeDan = "Nedelja"
    lblStatus.Caption = " Vreme: " + vreme + " | Datum: " + datum + " | Danas je: " + DanasJeDan + " | Tecomatic.rs"
End Sub
Public Sub lockApp(unlocked As Boolean)
If Not DebugMode = True Then On Error Resume Next
    If unlocked = True Then mnuzastitaz.Caption = "Zakljucaj"
    If unlocked = False Then mnuzastitaz.Caption = "Otkljucaj"
    mnuRaspored.Enabled = unlocked
    mnupodesavanja.Enabled = unlocked
    mnuinterfejs.Enabled = unlocked
    mnuizlaz.Enabled = unlocked
    mnuruucno.Enabled = unlocked
    CmbRaspored.Enabled = unlocked
    cmdPrimeni.Enabled = unlocked
End Sub
Public Sub PostaviZastitu()
If Not DebugMode = True Then On Error Resume Next
    If GetProfile("config", "035", "1", getConfigPath) = "0" Then
        mnuzastitaz.Visible = False
        lockApp True
    Else
        mnuzastitaz.Visible = True
        lockApp False
    End If
End Sub
Public Sub PostaviGridMain(ImeRasporeda As String)
If Not DebugMode = True Then On Error Resume Next
    Set tabela = New ADODB.Recordset
    tabela.Open "SELECT * FROM Raspored WHERE Raspored='" & CmbRaspored.Text & "'", Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    Set Grid.DataSource = tabela
    Grid.Refresh
    Grid.MarqueeStyle = dbgHighlightRow
    Grid.Columns.Remove ("ID")
    Grid.Columns.Remove ("Raspored")
    Grid.Columns.Remove ("Dan")
    Grid.Columns("Naziv").Width = 170
    Grid.Columns("Vreme").Width = 60
    Grid.Columns(2).Caption = "Duzina Zvona"
    Grid.Columns(2).Width = 80
    Grid.Columns(2).Alignment = dbgCenter
    If tabela.EOF = True Then
        WriteProfile "config", "021", CmbRaspored.Text, getConfigPath
    End If
    CmbRaspored.Text = ImeRasporeda
End Sub
Public Sub PostaviCombo()
If Not DebugMode = True Then On Error Resume Next
    Dim x As Integer
    CmbRaspored.Clear
    Set raspored = New ADODB.Recordset
    raspored.Open "SELECT DISTINCT Naziv FROM NaziviRasporeda", Me.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    For x = 1 To raspored.RecordCount
        CmbRaspored.AddItem raspored.Fields("Naziv").value
        raspored.MoveNext
    Next x
    raspored.Close
    CmbRaspored.Text = GetProfile("config", "021", "", getConfigPath)
End Sub
Private Sub CitajXml(Putanja As String)
If Not DebugMode = True Then On Error Resume Next
    Set uvoz = New ADODB.Recordset
    Dim NazivRasporeda As String, Poc As Byte, Kraj As Byte, Str As String, Str1 As String, i As Integer, j As Byte
    If Dir(Putanja) = "" Then Exit Sub
    Open Putanja For Input As #1
    i = 0
    Line Input #1, Str
    If Str <> "<Export>" Then
        MsgBox "Fajl je nepravilno formatiran!", vbCritical
        Close #1
        Exit Sub
    End If
    uvoz.Open "SELECT * FROM Raspored", Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    Do
    i = i + 1
    Line Input #1, Str
    If Trim(Str) = "</Export>" Then
        Close #1
        NazivRasporeda = uvoz.Fields("Raspored").value
        uvoz.Close
        If Not NazivRasporeda = "VanNastavni" Then
            uvoz.Open "SELECT * FROM NaziviRasporeda", Konekcija, adOpenStatic, adLockOptimistic, adCmdText
            uvoz.AddNew
            uvoz.Fields("Naziv").value = NazivRasporeda
            uvoz.Update
            uvoz.Close
            Call PostaviCombo
        End If
        CmbRaspored.Text = GetProfile("config", "021", "", getConfigPath)
        MsgBox "Raspored " & NazivRasporeda & " uspesno je uvezen u program!", vbInformation
        Exit Sub
    End If
    Poc = InStr(1, Str, "'")
    Kraj = InStr(Poc + 1, Str, "'")
    Str1 = Mid(Str, Poc + 1, Kraj - Poc - 1)
    
    With uvoz
    For i = 3 To 8
        Line Input #1, Str
        Poc = InStr(1, Str, ">")
        Kraj = InStr(Poc + 1, Str, "<")
        Str1 = Mid(Str, Poc + 1, Kraj - Poc - 1)
        If i = 4 Then .AddNew
        If i = 4 Then .Fields("Raspored").value = Str1
        If i = 5 Then .Fields("Naziv").value = Str1
        If i = 6 Then .Fields("Vreme").value = Str1
        If i = 7 Then .Fields("Dan").value = Str1
        If i = 8 Then .Fields("DuzinaZvona").value = Str1
    Next i
    .Update
    End With
    Line Input #1, Str
    DoEvents
    Loop Until EOF(1)
End Sub
Private Sub setInterfacePic()
On Error Resume Next
    If PortState = True Then
        If CheckInterface = True Then
            Me.ImgDissconn.Visible = False
            Me.ImgConnected.Visible = True
        Else
            Me.ImgDissconn.Visible = True
            Me.ImgConnected.Visible = False
        End If
    Else
        Me.ImgDissconn.Visible = False
        Me.ImgConnected.Visible = False
    End If
End Sub

