Attribute VB_Name = "MdlLanguage"
Option Explicit

Public CurrentLang As String

Private Type LangEntry
    section As String
    key     As String
    value   As String
End Type

Private langCache()  As LangEntry
Private langCount    As Long
Private langLoaded   As Boolean

Public Function GetSavedLang() As String
    Dim saved As String
    saved = GetProfile("config", "042", "Serbian", getConfigPath)
    If saved = "" Then saved = "Serbian"
    GetSavedLang = saved
End Function

Public Sub SaveLang(langCode As String)
    WriteProfile "config", "042", langCode, getConfigPath
End Sub

Public Function GetAvailableLanguages() As String()
    Dim result() As String
    Dim count As Integer
    count = 0
    ReDim result(0)
    Dim langPath As String
    langPath = App.Path & "\Language\"
    Dim fname As String
    fname = Dir(langPath & "*_language.lng")
    Do While fname <> ""
        Dim code As String
        code = Left(fname, InStr(fname, "_") - 1)
        If count > 0 Then ReDim Preserve result(count)
        result(count) = code
        count = count + 1
        fname = Dir()
    Loop
    If count = 0 Then
        result(0) = "sr"
    End If
    GetAvailableLanguages = result
End Function

Public Sub InitLanguage(langCode As String)
    If langCode = "" Then langCode = "Serbian"
    CurrentLang = langCode
    langLoaded = False
    langCount = 0
    ReDim langCache(63)
    Call LoadLangFile(App.Path & "\Language\" & langCode & "_language.lng")
End Sub

Public Function t(section As String, key As String) As String
    If Not langLoaded Then
        t = key
        Exit Function
    End If
    Dim i As Long
    Dim sU As String: sU = UCase(section)
    Dim kU As String: kU = UCase(key)
    For i = 0 To langCount - 1
        If UCase(langCache(i).section) = sU And UCase(langCache(i).key) = kU Then
            t = langCache(i).value
            Exit Function
        End If
    Next i
    t = key
End Function

Public Sub ApplyLanguage(frm As Form)
    If Not langLoaded Then Exit Sub

    Dim sName As String
    sName = frm.Name

    Dim cap As String
    cap = t(sName, "Caption")
    If cap <> "Caption" Then frm.Caption = cap

    On Error Resume Next

    Dim ctrl As Control
    For Each ctrl In frm.Controls
        Dim ctrlName As String
        ctrlName = ctrl.Name
        Dim val As String
        val = t(sName, ctrlName)
        If val <> ctrlName Then
            ctrl.Caption = val
        End If
        Dim tipVal As String
        tipVal = t(sName, ctrlName & "_Tip")
        If tipVal <> ctrlName & "_Tip" Then
            ctrl.ToolTipText = tipVal
        End If
    Next ctrl

    On Error GoTo 0
End Sub

Public Sub ApplyLanguageToAll()
    Dim frm As Form
    For Each frm In Forms
        ApplyLanguage frm
    Next frm
End Sub
Public Sub RefreshAllForms()
    Call ApplyLanguageToAll
    On Error Resume Next
    If Not FrmMain Is Nothing Then
        If FrmMain.Visible Then
            Call FrmMain.RefreshLanguage
        End If
    End If
    On Error GoTo 0
End Sub

Public Function GetDayName(index As Integer) As String
    Select Case index
        Case 1: GetDayName = t("FrmVanNastavne", "ItemPonedeljak")
        Case 2: GetDayName = t("FrmVanNastavne", "ItemUtorak")
        Case 3: GetDayName = t("FrmVanNastavne", "ItemSreda")
        Case 4: GetDayName = t("FrmVanNastavne", "ItemCetvrtak")
        Case 5: GetDayName = t("FrmVanNastavne", "ItemPetak")
        Case 6: GetDayName = t("FrmVanNastavne", "ItemSubota")
        Case 7: GetDayName = t("FrmVanNastavne", "ItemNedelja")
        Case Else: GetDayName = ""
    End Select
End Function

Public Function GetDayIndexByName(dayName As String) As Integer
    Dim i As Integer
    For i = 1 To 7
        If GetDayName(i) = dayName Then
            GetDayIndexByName = i
            Exit Function
        End If
    Next i
    Select Case dayName
        Case "Ponedeljak": GetDayIndexByName = 1
        Case "Utorak":     GetDayIndexByName = 2
        Case "Sreda":      GetDayIndexByName = 3
        Case "Cetvrtak":   GetDayIndexByName = 4
        Case "Petak":      GetDayIndexByName = 5
        Case "Subota":     GetDayIndexByName = 6
        Case "Nedelja":    GetDayIndexByName = 7
        Case Else:         GetDayIndexByName = 0
    End Select
End Function

Private Sub LoadLangFile(filePath As String)
    langLoaded = False
    langCount = 0
    ReDim langCache(63)

    Dim fNum As Integer
    fNum = FreeFile

    On Error GoTo FileError
    Open filePath For Input As #fNum
    On Error GoTo 0

    Dim currentSection As String
    Dim lineStr As String
    Dim eqPos As Integer

    Do While Not EOF(fNum)
        Line Input #fNum, lineStr
        lineStr = Trim(lineStr)
        If Len(lineStr) = 0 Then GoTo SkipLine
        If Left(lineStr, 1) = ";" Then GoTo SkipLine
        If Left(lineStr, 1) = "[" Then
            Dim rb As Integer
            rb = InStr(lineStr, "]")
            If rb > 1 Then currentSection = Mid(lineStr, 2, rb - 2)
            GoTo SkipLine
        End If
        eqPos = InStr(lineStr, "=")
        If eqPos > 1 And Len(currentSection) > 0 Then
            If langCount > UBound(langCache) Then
                ReDim Preserve langCache(langCount + 63)
            End If
            langCache(langCount).section = currentSection
            langCache(langCount).key = Trim(Left(lineStr, eqPos - 1))
            langCache(langCount).value = Mid(lineStr, eqPos + 1)
            langCount = langCount + 1
        End If
SkipLine:
    Loop
    Close #fNum
    langLoaded = True
    Exit Sub

FileError:
    langLoaded = False
End Sub
