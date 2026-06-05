VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.UserControl CntrlSerial 
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4125
   ScaleHeight     =   2385
   ScaleWidth      =   4125
   Begin MSCommLib.MSComm COM 
      Left            =   960
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer trm 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   2520
      Top             =   480
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   615
      Left            =   0
      Picture         =   "CntrlSeria1l.ctx":0000
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
                .DTREnable = False
            End With
            If COM.PortOpen = True Then
                PrikaziPoruku "Port: " & portNumber & " je otvoren!", "5"
                OpenPort = True
                Exit Function
            End If
        End If
End Function
Public Function ClosePort() As Boolean
If Not DebugMode = True Then On Error Resume Next
    If COM.PortOpen = True Then COM.PortOpen = False
    PrikaziPoruku "Port zatvoren! " & PortBroj, "5"
End Function
Public Function CheckInterface() As Boolean
If Not DebugMode = True Then On Error Resume Next
    If COM.PortOpen = True Then CheckInterface = True
End Function
Public Function PortState() As Boolean
If Not DebugMode = True Then On Error Resume Next
    PortState = COM.PortOpen
End Function
Public Sub HitTheRelay(status As Boolean)
If Not DebugMode = True Then On Error Resume Next
If status = True Then
    status = False
Else
    status = True
End If

    If GetProfile("config", "024", "", getConfigPath) = "Bell USB" Then Trm.Enabled = status
    If GetProfile("config", "024", "", getConfigPath) = "Bell Ethernet" Then
        'ethernet interfejs
    End If
    If GetProfile("config", "024", "", getConfigPath) = "Bell Comm" Then
        If COM.PortOpen = True Then COM.RTSEnable = status
    End If
    If GetProfile("config", "024", "", getConfigPath) = "Bez interfejsa" Then
        Exit Sub
    End If
End Sub

Private Sub trm_Timer()
If Not DebugMode = True Then On Error Resume Next

DoEvents
    If DebugMode = False Then
        If CheckInterface = True Then COM.output = "                   "
    Else
        COM.output = "                   "
    End If
End Sub

