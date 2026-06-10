Attribute VB_Name = "MdlBellProRelay"
Option Explicit

'========================
' Windows API
'========================
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

'========================
' DEVICE HANDLE
'========================
Private hDevice As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3

'========================
' OPEN DEVICE
'========================
Public Function BellProRelay_Open() As Boolean

    hDevice = CreateFile("\\.\HID#VID_16D0&PID_0753", _
                         GENERIC_READ Or GENERIC_WRITE, _
                         0, 0, OPEN_EXISTING, 0, 0)

    BellProRelay_Open = (hDevice <> -1)

End Function

'========================
' SEND 1 BYTE COMMAND
'========================
Public Sub BellProRelay_Send(ByVal cmd As Byte)

    Dim buffer(0) As Byte
    Dim written As Long

    buffer(0) = cmd

    WriteFile hDevice, buffer(0), 1, written, 0

End Sub

'========================
' READ RESPONSE
'========================
Public Function BellProRelay_Read() As Byte

    Dim buffer(0) As Byte
    Dim read As Long

    ReadFile hDevice, buffer(0), 1, read, 0

    If read > 0 Then
        BellProRelay_Read = buffer(0)
    Else
        BellProRelay_Read = &HFF
    End If

End Function

'========================
' PING CHECK
'========================
Public Function BellProRelay_IsAlive() As Boolean

    Dim resp As Byte
    Dim t As Single

    BellProRelay_Send &H1

    t = Timer

    Do
        resp = BellProRelay_Read()

        If resp = &HAA Then
            BellProRelay_IsAlive = True
            Exit Function
        End If

        DoEvents

        If Timer - t > 1 Then Exit Do

    Loop

    BellProRelay_IsAlive = False

End Function

'========================
' HIGH LEVEL API
'========================
Public Sub BellProRelay_Relay1On()
    BellProRelay_Send &H11
End Sub

Public Sub BellProRelay_Relay1Off()
    BellProRelay_Send &H10
End Sub

Public Sub BellProRelay_Relay2On()
    BellProRelay_Send &H21
End Sub

Public Sub BellProRelay_Relay2Off()
    BellProRelay_Send &H20
End Sub

Public Function BellProRelay_GetState() As Byte
    BellProRelay_Send &H30
    BellProRelay_GetState = BellProRelay_Read()
End Function

'========================
' CLOSE
'========================
Public Sub BellProRelay_Close()

    If hDevice <> 0 Then
        CloseHandle hDevice
        hDevice = 0
    End If

End Sub

