Attribute VB_Name = "MdlBellLogic"
Option Explicit
Public Function ProveriSve() As Boolean
If Not DebugMode = True Then On Error Resume Next

If GetProfile("config", "030", "0", getConfigPath) = 1 Then
    If Subota() = True Then
        ProveriSve = True
        Exit Function
    Else
        ProveriSve = False
    End If
End If

If GetProfile("config", "031", "0", getConfigPath) = 1 Then
    If Nedelja() = True Then
        ProveriSve = True
        Exit Function
    Else
        ProveriSve = False
    End If
End If

If GetProfile("config", "034", "0", getConfigPath) = 1 Then
    If Praznik() = True Then
        ProveriSve = True
        Exit Function
    Else
        ProveriSve = False
    End If
End If

If GetProfile("config", "032", "0", getConfigPath) = 1 Then
    If Raspust() = True Then
        ProveriSve = True
        Exit Function
    Else
        ProveriSve = False
    End If
End If

If GetProfile("config", "028", "0", getConfigPath) = 1 Then
    ProveriSve = True
    Exit Function
End If
ProveriSve = False
End Function
Private Function Praznik() As Boolean
If Not DebugMode = True Then On Error Resume Next
    Dim x As Integer
        For x = 1 To 12
            If ReadConfig(format(x, "000")) = FrmMain.datum Then
                Praznik = True
                Exit Function
            End If
        Next x
End Function
Private Function Raspust() As Boolean
If Not DebugMode = True Then On Error Resume Next
Dim dan As String, mesec As String, godina As String
dan = format(Now, "dd")
mesec = format(Now, "MM")
godina = format(Now, "yyyy")

    If godina >= GetRightFormat("yyyy", ReadConfig("013")) And godina <= GetRightFormat("yyyy", ReadConfig("014")) Then
        If mesec >= GetRightFormat("mm", ReadConfig("013")) And mesec <= GetRightFormat("mm", ReadConfig("014")) Then
            If dan >= GetRightFormat("dd", ReadConfig("013")) And dan <= GetRightFormat("dd", ReadConfig("014")) Then
                Raspust = True
                Exit Function
            Else
                Raspust = False
            End If
        Else
            Raspust = False
        End If
    Else
        Raspust = False
    End If
    
    If godina >= GetRightFormat("yyyy", ReadConfig("015")) And godina <= GetRightFormat("yyyy", ReadConfig("016")) Then
        If mesec >= GetRightFormat("mm", ReadConfig("015")) And mesec <= GetRightFormat("mm", ReadConfig("016")) Then
            If dan >= GetRightFormat("dd", ReadConfig("015")) And dan <= GetRightFormat("dd", ReadConfig("016")) Then
                Raspust = True
                Exit Function
            Else
                Raspust = False
            End If
        Else
            Raspust = False
        End If
    Else
        Raspust = False
    End If
    
    If godina >= GetRightFormat("yyyy", ReadConfig("017")) And godina <= GetRightFormat("yyyy", ReadConfig("018")) Then
        If mesec >= GetRightFormat("mm", ReadConfig("017")) And mesec <= GetRightFormat("mm", ReadConfig("018")) Then
            If dan >= GetRightFormat("dd", ReadConfig("017")) And dan <= GetRightFormat("dd", ReadConfig("018")) Then
                Raspust = True
                Exit Function
            Else
                Raspust = False
            End If
        Else
            Raspust = False
        End If
    Else
     Raspust = False
    End If
    
    If godina >= GetRightFormat("yyyy", ReadConfig("019")) And godina <= GetRightFormat("yyyy", ReadConfig("020")) Then
        If mesec >= GetRightFormat("mm", ReadConfig("019")) And mesec <= GetRightFormat("mm", ReadConfig("020")) Then
            If dan >= GetRightFormat("dd", ReadConfig("019")) And dan <= GetRightFormat("dd", ReadConfig("020")) Then
                Raspust = True
                Exit Function
            Else
                Raspust = False
            End If
        Else
            Raspust = False
        End If
    Else
        Raspust = False
    End If
End Function
Private Function Subota() As Boolean
If Not DebugMode = True Then On Error Resume Next
If Weekday(Date) = 7 Then
    Subota = True
Else
    Subota = False
End If
End Function
Private Function Nedelja() As Boolean
If Not DebugMode = True Then On Error Resume Next
    If Weekday(Date) = 1 Then
        Nedelja = True
    Else
        Nedelja = False
    End If
End Function
