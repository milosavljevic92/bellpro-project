Attribute VB_Name = "MdlBellProRelay"
Option Explicit

'==============================
' BellProRelay USB Interface
' Tecomatic - Novi Sad
'==============================

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

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private hDevice As Long

Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const OPEN_EXISTING As Long = 3

'==============================
' OPEN DEVICE
'==============================
Public Function BellProRelay_Open() As Boolean

    ' DigiUSB / ATtiny85 HID device (Digispark default VID/PID)
    hDevice = CreateFile("\\.\HID#VID_16D0&PID_0753",
                         GENERIC_READ Or GENERIC_WRITE,
                         0, 0, OPEN_EXISTING, 0, 0)

    BellProRelay_Open = (hDevice <> -1)

End Function

'==============================
' SEND COMMAND
'==============================
Public Sub BellProRelay_Send(ByVal cmd As String)

    Dim buffer(63) As Byte
    Dim i As Integer
    Dim written As Long

    For i = 1 To Len(cmd)
        buffer(i - 1) = Asc(Mid$(cmd, i, 1))
    Next i

    buffer(Len(cmd)) = 10 ' LF

    WriteFile hDevice, buffer(0), 64, written, 0

End Sub

'==============================
' CLOSE DEVICE
'==============================
Public Sub BellProRelay_Close()

    If hDevice <> 0 Then
        CloseHandle hDevice
        hDevice = 0
    End If

End Sub
