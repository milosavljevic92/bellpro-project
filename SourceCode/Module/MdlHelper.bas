Attribute VB_Name = "MdlHelper"
Option Explicit
Private Declare Function GetVolumeSerialNumber Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public DebugMode As Boolean
Public AboutApp As Boolean
Private tmp As Boolean
Public Function VolumeSerialNumber(ByVal RootPath As String) As String
    If Not DebugMode = True Then On Error Resume Next
    Dim VolLabel As String, VolSize As Long, Serial As Long, MaxLen As Long, flags As Long, Name As String, NameSize As Long, s As String, ret As Boolean
    ret = GetVolumeSerialNumber(RootPath, VolLabel, VolSize, Serial, MaxLen, flags, Name, NameSize)
    s = format(Hex(Serial), "00000000")
    VolumeSerialNumber = Left(s, 4) & Right(s, 4)
End Function
Public Function ReturnNonAlpha(ByVal sString As String) As String
   Dim i As Integer
   For i = 1 To Len(sString)
       If Mid(sString, i, 1) Like "[0-9]" Then
           ReturnNonAlpha = ReturnNonAlpha + Mid(sString, i, 1)
       End If
   Next i
End Function
Public Function checkDoesDBexist(Path As String) As Boolean
    Dim fsoBase  As New Scripting.FileSystemObject
    If fsoBase.FileExists(Path) = False Then
        checkDoesDBexist = False
    Else
        checkDoesDBexist = True
    End If
End Function
Public Sub setDemoMode()
    tmp = True
    Exit Sub
End Sub
Public Sub setFullMode()
    tmp = False
    Exit Sub
End Sub
Public Function ProveriFormatVremena(vreme As String) As Boolean
If Not DebugMode = True Then On Error Resume Next
    If Len(vreme) < 5 Or Mid(vreme, 3, 1) <> ":" Or Mid(vreme, 1, 2) > 23 Or Mid(vreme, 4, 2) > 59 Then
        ProveriFormatVremena = False
    Else
        ProveriFormatVremena = True
    End If
End Function
Public Function getDemoStatus() As Boolean
    getDemoStatus = tmp
    Exit Function
End Function
Public Function getConfigPath() As String
    getConfigPath = App.Path + "\config.ini"
End Function
Public Sub setDebugMode(isOn As Boolean)
    DebugMode = isOn
End Sub
Public Sub setAboutApp(isOn As Boolean)
    AboutApp = isOn
End Sub
Public Function GetRightFormat(formatPath As String, dateToConvert As String) As String
    GetRightFormat = format(dateToConvert, formatPath)
End Function
Public Function ReadConfig(section As String) As String
    ReadConfig = GetProfile("config", section, "", getConfigPath)
End Function
Public Sub PrikaziPoruku(poruka As String, Duzina As String)
    Dim prozor As Form
    Set prozor = New FrmZvoni
    prozor.Visible = False
    prozor.VremeZatvaranja Duzina
    prozor.LblText.Caption = "Tecomatic - BellPro"
    prozor.LblMessage.Caption = Replace(poruka, ", ", vbCrLf & vbCrLf)
    prozor.Visible = True
End Sub
Public Sub GenerateIniFileIfNotExist()
    Dim fs
    Dim fsoIni As New Scripting.FileSystemObject
    If fsoIni.FileExists(getConfigPath) = False Then
        Set fs = fsoIni.CreateTextFile(getConfigPath)
        fs.Write _
        "[config]" & vbCrLf & _
        "001=" & vbCrLf & _
        "002=" & vbCrLf & _
        "003=" & vbCrLf & _
        "004=" & vbCrLf & _
        "005=" & vbCrLf & _
        "006=" & vbCrLf & _
        "007=" & vbCrLf & _
        "008=" & vbCrLf & _
        "009=" & vbCrLf & _
        "010=" & vbCrLf & _
        "011=" & vbCrLf & _
        "012=" & vbCrLf & _
        "013=" & GenerateTodayDate & vbCrLf & _
        "014=" & GenerateTodayDate & vbCrLf & _
        "015=" & GenerateTodayDate & vbCrLf & _
        "016=" & GenerateTodayDate & vbCrLf & _
        "017=" & GenerateTodayDate & vbCrLf & _
        "018=" & GenerateTodayDate & vbCrLf & _
        "019=" & GenerateTodayDate & vbCrLf & _
        "020=" & GenerateTodayDate & vbCrLf & _
        "021=" & "Svakodnevni" & vbCrLf & _
        "022=" & vbCrLf
        fs.Write _
        "023=" & "COM1" & vbCrLf & _
        "024=" & "Bell USB" & vbCrLf & _
        "025=" & "12" & vbCrLf & _
        "026=" & "1" & vbCrLf & _
        "027=" & "1" & vbCrLf & _
        "028=" & "0" & vbCrLf & _
        "029=" & "1" & vbCrLf & _
        "030=" & "1" & vbCrLf & _
        "031=" & "1" & vbCrLf & _
        "032=" & "0" & vbCrLf & _
        "033=" & vbCrLf & _
        "034=" & "0" & vbCrLf & _
        "035=" & "0" & vbCrLf & _
        "036=" & "0" & vbCrLf & _
        "037=" & "0" & vbCrLf
        fs.Close
    End If
End Sub
Private Function GenerateTodayDate() As String
GenerateTodayDate = format(Now, "dd.MM.yyyy")
End Function
Public Function Kriptuj(txtString As String, EnCrypt As Boolean) As String
On Error Resume Next
Dim x As Integer, outString As String, iLen As Integer, sFirstSeed As String, sSecondSeed As String, iSeed As Integer
       If EnCrypt Then
           sFirstSeed = Left(txtString, 1)
           sSecondSeed = Mid(txtString, 2, 1)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString)
           For x = 1 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) + 2) & outString
           Next
           outString = Chr(Asc(sFirstSeed) * 2 + 3) & outString
           outString = outString & Chr(Asc(sSecondSeed) * 2 - 3)
       Else
           sFirstSeed = Chr((Asc(Left(txtString, 1)) - 3) \ 2)
           sSecondSeed = Chr((Asc(Right(txtString, 1)) + 3) \ 2)
           iSeed = (Asc(sFirstSeed) + Asc(sSecondSeed)) Mod 2
           iLen = Len(txtString) - 1
           For x = 2 To iLen
               outString = Chr((Asc(Mid$(txtString, x, 1)) Xor iSeed) - 2) & outString
           Next
        End If
       Kriptuj = outString
       Exit Function
End Function
