Attribute VB_Name = "MdlStart"
Option Explicit
Public Sub Main()
On Error Resume Next
    If App.PrevInstance = True Then
        MsgBox "Program je veæ pokrenut!", vbExclamation
        End
    End If
    GenerateIniFileIfNotExist
    setDebugMode False 'debug mod
    If GetProfile("config", "037", "0", getConfigPath) = "1" Then initPlayer
    'setLang
    FrmRegistracija.Show
    Exit Sub
End Sub
