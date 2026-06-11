Attribute VB_Name = "MdlDatabase"
Public Function CreateDatabaseIfNotExists() As Boolean

    On Error GoTo Greska

    Dim cn As ADODB.Connection
    Dim dbFile As String

    dbFile = App.Path & "\BellPro.sqlite"

    If Dir$(dbFile) <> "" Then
        CreateDatabaseIfNotExists = True
        Exit Function
    End If

    Set cn = New ADODB.Connection
    cn.Open "Driver={SQLite3 ODBC Driver};Database=" & dbFile & ";"
    cn.Execute _
        "CREATE TABLE NaziviRasporeda (" & _
        "ID INTEGER PRIMARY KEY, " & _
        "Naziv TEXT)"

    cn.Execute _
        "CREATE TABLE Raspored (" & _
        "ID INTEGER PRIMARY KEY AUTOINCREMENT, " & _
        "Raspored TEXT, " & _
        "Naziv TEXT, " & _
        "Vreme TEXT, " & _
        "Dan TEXT, " & _
        "DuzinaZvona TEXT)"

    cn.Close
    Set cn = Nothing

    CreateDatabaseIfNotExists = True
    Exit Function

Greska:

    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
    End If

    Set cn = Nothing

    CreateDatabaseIfNotExists = False

End Function

