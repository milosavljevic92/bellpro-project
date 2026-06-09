VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.UserControl CntrlSerial 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   780
   ScaleHeight     =   1605
   ScaleWidth      =   780
   Begin MSCommLib.MSComm COM 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      Picture         =   "CntrlSerial.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "CntrlSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function OpenPort(portNumber As String) As Boolean
If Not DebugMode = True Then On Error Resume Next
        Dim portNumberInt As Integer
        portNumberInt = Mid(portNumber, 4, 2)
        If IsCommExist(portNumberInt) = False Then
            PrikaziPoruku "Greska sa otvaranjem porta: " & PortBroj, "5"
            OpenPort = False
            Exit Function
        Else
            With COM
                .CommPort = Mid(portNumber, 4, 2)
                .Settings = "2400,n,8,1"
                .PortOpen = True
            
            End With
            
            If GetProfile("config", "040", "", getConfigPath) = "1" Then
                COM.DTREnable = True
            Else
                COM.DTREnable = False
            End If
            
            If GetProfile("config", "041", "", getConfigPath) = "1" Then
                COM.RTSEnable = True
            Else
                COM.RTSEnable = False
            End If
            
            If COM.PortOpen = True Then
                PrikaziPoruku "Port: " & portNumber & " je otvoren!", "5"
                OpenPort = True
                Exit Function
            End If
        End If
End Function
Public Function ClosePort() As Boolean
If Not DebugMode = True Then On Error Resume Next
    If GetProfile("config", "041", "", getConfigPath) = "1" Then
        COM.RTSEnable = True
    Else
        COM.RTSEnable = False
    End If
    If COM.PortOpen = True Then COM.PortOpen = False
    PrikaziPoruku "Port zatvoren! " & PortBroj, "5"
End Function
Public Function CheckInterface() As Boolean
If Not DebugMode = True Then On Error Resume Next
    If GetProfile("config", "040", "", getConfigPath) = "1" Then
        If COM.PortOpen = True Then
            COM.DTREnable = True
            If COM.CTSHolding = True Then
                CheckInterface = True
            Else
                CheckInterface = False
            End If
        Else
            CheckInterface = False
        End If
    Else
        If COM.PortOpen = True Then CheckInterface = True
    End If
End Function
Public Function PortState() As Boolean
If Not DebugMode = True Then On Error Resume Next
    PortState = COM.PortOpen
End Function
Public Sub HitTheRelay(status As Boolean)
If Not DebugMode = True Then On Error Resume Next
    If GetProfile("config", "024", "", getConfigPath) = "Bell Comm" Then
        If CheckInterface() = True Then
            If GetProfile("config", "041", "", getConfigPath) = "1" Then
                COM.RTSEnable = Not status
            Else
                COM.RTSEnable = status
            End If
        End If
    End If
            
    If GetProfile("config", "024", "", getConfigPath) = "Bez interfejsa" Then
        Exit Sub
    End If
End Sub




