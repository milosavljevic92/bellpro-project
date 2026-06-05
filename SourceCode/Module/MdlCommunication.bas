Attribute VB_Name = "MdlCommunication"
Option Explicit
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Function OpenPort() As Boolean
On Error Resume Next
    If getDemoStatus = True Then
        OpenPort = False
        Exit Function
    End If
    Dim portNumber As String
    portNumber = GetProfile("config", "023", "COM1", getConfigPath)
    OpenPort = FrmMain.CntrlSerial.OpenPort(portNumber)
    If OpenPort = True Then
        FrmMain.ImgDissconn.ToolTipText = "Port otvoren: " + portNumber
    Else
        FrmMain.ImgDissconn.ToolTipText = "Greska sa otvaranjem porta!"
    End If
End Function
Public Function ClosePort() As Boolean
    ClosePort = FrmMain.CntrlSerial.ClosePort
    PrikaziPoruku "Port zatvoren! ", "5"
End Function
Public Function CheckInterface() As Boolean
    CheckInterface = FrmMain.CntrlSerial.CheckInterface
End Function
Public Function PortState() As Boolean
    PortState = FrmMain.CntrlSerial.PortState
End Function
Public Sub HitTheRelay(status As Boolean)
    If getDemoStatus = True Then
        Exit Sub
    End If
    FrmMain.CntrlSerial.HitTheRelay (status)
End Sub
Public Function IsCommExist(COMNum As Integer) As Boolean
If Not DebugMode = True Then On Error Resume Next
    Dim hCOM As Long
    Dim ret As Long
    Dim sec As SECURITY_ATTRIBUTES
    hCOM = CreateFile("\.\COM" & COMNum & "", 0, FILE_SHARE_READ + _
        FILE_SHARE_WRITE, sec, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If hCOM = -1 Then
        IsCommExist = False
    Else
        IsCommExist = True
        ret = CloseHandle(hCOM)
    End If
End Function

 


