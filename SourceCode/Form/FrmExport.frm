VERSION 5.00
Begin VB.Form FrmExport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Izvoz rasporeda:"
   ClientHeight    =   960
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   Icon            =   "FrmExport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   960
   ScaleWidth      =   5655
   StartUpPosition =   2  'CenterScreen
   Begin BellPro.XPButton cmdIzvezi 
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
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
      Caption         =   "Izvezi"
      ForeColor       =   -2147483642
      ForeHover       =   0
   End
   Begin VB.ComboBox CmbRaspored 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Label LblRaspored 
      BackStyle       =   0  'Transparent
      Caption         =   "Raspored:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "FrmExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private raspored As ADODB.Recordset
Private tabela As ADODB.Recordset
Private Sub CmdIzvezi_Click()
If Not DebugMode = True Then On Error Resume Next
    Dim SaveFile As SAVEFILENAME
    Dim lReturn As Long
    Dim sFilter As String
    SaveFile.lStructSize = Len(SaveFile)
    SaveFile.hwndOwner = FrmExport.hwnd
    SaveFile.hInstance = App.hInstance
    sFilter = "eXtensible Markup Language (*.xml)" & Chr(0) & "*.XML" & Chr(0)
    SaveFile.lpstrFilter = sFilter
    SaveFile.nFilterIndex = 1
    SaveFile.lpstrFile = String(257, 0)
    SaveFile.nMaxFile = Len(SaveFile.lpstrFile) - 1
    SaveFile.lpstrFileTitle = SaveFile.lpstrFile
    SaveFile.nMaxFileTitle = SaveFile.nMaxFile
    SaveFile.lpstrInitialDir = "C:\"
    SaveFile.lpstrTitle = "Sacuvaj XML u ..."
    SaveFile.flags = 0
    lReturn = GetSaveFileName(SaveFile)
    If lReturn = 0 Then
        Exit Sub
    Else
        If CmbRaspored.Text = "Vannastavne aktivnosti" Then
                GenerisiXML Trim(SaveFile.lpstrFile) & ".xml", "SELECT * FROM raspored WHERE Raspored='VanNastavni'"
            Else
                GenerisiXML Trim(SaveFile.lpstrFile) & ".xml", "SELECT * FROM raspored WHERE Raspored='" & CmbRaspored.Text & "'"
            End If
    End If
End Sub
Private Sub Form_Load()
If Not DebugMode = True Then On Error Resume Next
    Call PostaviComboExport
End Sub
Private Sub PostaviComboExport()
If Not DebugMode = True Then On Error Resume Next
    Dim x As Integer
    CmbRaspored.Clear
    Set raspored = New ADODB.Recordset
    raspored.Open "SELECT DISTINCT Naziv FROM NaziviRasporeda", FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
    CmbRaspored.AddItem "Vannastavne aktivnosti"
    For x = 1 To raspored.RecordCount
       CmbRaspored.AddItem raspored.Fields("Naziv").value
       raspored.MoveNext
    Next x
    raspored.Close
    CmbRaspored.Text = "Vannastavne aktivnosti"
End Sub
Private Sub GenerisiXML(ImeFajla As String, sqlString As String)
If Not DebugMode = True Then On Error Resume Next
    Set tabela = New ADODB.Recordset
    Dim tRecs As Double
    Dim x As Integer, i As Integer
    With tabela
            .Open sqlString, FrmMain.Konekcija, adOpenStatic, adLockOptimistic, adCmdText
            If .RecordCount = 0 Then
                .Close
                MsgBox "Raspored je prazan i ne moze biti izvezen", vbCritical
                Exit Sub
            End If
            Open ImeFajla For Output As #1
            Print #1, "<" & "Export" & ">"
            .MoveFirst
            Do Until .EOF = True
            DoEvents
            tRecs = tRecs + 1
            .MoveNext
            Loop
            .MoveFirst
            If .RecordCount <> 0 Then
                Do Until .EOF = True
                x = x + 1
                DoEvents
                Print #1, "<record exported='" & Now & "' id='" & x & "'>"
                For i = 0 To .Fields.Count - 1
                DoEvents
                Print #1, "<" & .Fields(i).Name & ">" & .Fields(i).value & "</" & .Fields(i).Name & ">"
                Next i
                Print #1, "</record>"
                .MoveNext
                Loop
                Print #1, "</" & "Export" & ">"
                Close #1
            Else
               Exit Sub
            End If
            .Close
    End With
           MsgBox "Uspesno izvezen raspored!", vbInformation
End Sub
Private Sub Form_Unload(Cancel As Integer)
If Not DebugMode = True Then On Error Resume Next
    FrmMain.Show
    Unload Me
End Sub
