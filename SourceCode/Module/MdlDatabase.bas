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

    ' Tabela rasporeda
    cn.Execute _
        "CREATE TABLE NaziviRasporeda (" & _
        "ID INTEGER PRIMARY KEY, " & _
        "Naziv TEXT)"

    ' Stavke rasporeda
    cn.Execute _
        "CREATE TABLE Raspored (" & _
        "ID INTEGER PRIMARY KEY AUTOINCREMENT, " & _
        "Raspored TEXT, " & _
        "Naziv TEXT, " & _
        "Vreme TEXT, " & _
        "Dan TEXT, " & _
        "DuzinaZvona INTEGER)"

    ' Podrazumevani raspored
    cn.Execute _
        "INSERT INTO NaziviRasporeda (ID, Naziv) " & _
        "VALUES (1, 'Svakodnevni')"

    ' Najava
    cn.Execute _
        "INSERT INTO Raspored " & _
        "(Raspored, Naziv, Vreme, Dan, DuzinaZvona) VALUES (" & _
        "'Svakodnevni'," & _
        "'Najava'," & _
        "'07:30:00'," & _
        "'pon - ned'," & _
        "3)"

    ' Poèetak prvog èasa
    cn.Execute _
        "INSERT INTO Raspored " & _
        "(Raspored, Naziv, Vreme, Dan, DuzinaZvona) VALUES (" & _
        "'Svakodnevni'," & _
        "'Pocetak prvog casa'," & _
        "'08:00:00'," & _
        "'pon - ned'," & _
        "5)"

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
