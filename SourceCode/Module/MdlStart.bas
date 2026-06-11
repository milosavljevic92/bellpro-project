Attribute VB_Name = "MdlStart"
Option Explicit
Public Sub Main()
On Error Resume Next
    If App.PrevInstance = True Then
        MsgBox t("MdlStart", "MsgVecPokrenut"), vbExclamation
        End
    End If
    GenerateIniFileIfNotExist
    setDebugMode False
    If GetProfile("config", "037", "0", getConfigPath) = "1" Then initPlayer
    Call InitLanguage(GetSavedLang())
    FrmRegistracija.Show
    Exit Sub
End Sub
