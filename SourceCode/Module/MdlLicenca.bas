Attribute VB_Name = "MdlLicenca"
Option Explicit

Public UndoTemp         As String

Private numConv(10)     As Byte
Private Plain           As Variant
Private Code(28)        As String
Private Square(36)      As String
Private SquareCode(5)   As String
Private Row()           As Integer
Private PlaySquare      As String

'----------------------------------------------------------------
'
'                       Ceasar Shift
'
'----------------------------------------------------------------

Public Function EncodeCeasar(ByVal PlainIn As String, ByVal key As String) As String
'encode with Single Columnar
Dim i As Long
Dim plainC As Integer
Dim codeC As Integer
Dim shiftC As Integer

'check key and text lenght
If Len(key) <> 1 Then
    MsgBox "The Shift Key must be one letter, representing the begin of the shifted row.", vbCritical
    Exit Function
    End If

'trim all but alphabet
PlainIn = TrimText(PlainIn, True, False, False, False)
If PlainIn = "" Then Exit Function

shiftC = Asc(key) - 65

'encode
For i = 1 To Len(PlainIn)
    plainC = Asc(Mid(PlainIn, i, 1)) - 64
    codeC = plainC + shiftC
    If codeC > 26 Then codeC = codeC - 26
    EncodeCeasar = EncodeCeasar & Chr(codeC + 64)
Next i

End Function


Public Function DecodeCeasar(ByVal CodeIn As String, ByVal key As String) As String
'decode with Single Columnar
Dim i As Long
Dim plainC As Integer
Dim codeC As Integer
Dim shiftC As Integer

'check key and text lenght
If Len(key) <> 1 Then
    MsgBox "The Shift Key must be one letter, representing the begin of the shifted row.", vbCritical
    Exit Function
    End If

'trim all but alphabet
CodeIn = TrimText(CodeIn, True, False, False, False)
If CodeIn = "" Then Exit Function

shiftC = Asc(key) - 65

'decode
For i = 1 To Len(CodeIn)
    codeC = Asc(Mid(CodeIn, i, 1)) - 64
    plainC = codeC - shiftC
    If plainC < 1 Then plainC = plainC + 26
    DecodeCeasar = DecodeCeasar & Chr(plainC + 64)
Next i

End Function

'----------------------------------------------------------------
'
'          Single and Double Columnar Transposition
'
'----------------------------------------------------------------

Public Function EncodeColumnar(ByVal PlainIn As String, ByVal key As String) As String
'encode with Single Columnar

'trim all but alphabet
PlainIn = TrimText(PlainIn, True, False, False, False)
If PlainIn = "" Then Exit Function

'initialize columnar key
key = TrimText(key, True, False, False, False)
If InitColumnar(key) <> 0 Then Exit Function

'encode
EncodeColumnar = EncColumn(PlainIn)

End Function


Public Function DecodeColumnar(ByVal CodeIn As String, ByVal key As String) As String
'decode with Single Columnar

CodeIn = TrimText(CodeIn, True, False, False, False)
If CodeIn = "" Then Exit Function

'initialize columnar key
key = TrimText(key, True, False, False, False)
If InitColumnar(key) <> 0 Then Exit Function

'decode
DecodeColumnar = DecColumn(CodeIn)

End Function


Public Function EncodeDoubleColumnar(ByVal PlainIn As String, ByVal keyCol1 As String, ByVal keyCol2 As String) As String
'encode with Double Columnar

'trim all but alphabet
PlainIn = TrimText(PlainIn, True, False, False, False)
If PlainIn = "" Then Exit Function

'initialize 1st columnar key
keyCol1 = TrimText(keyCol1, True, False, False, False)
If InitColumnar(keyCol1) <> 0 Then Exit Function

'encode
EncodeDoubleColumnar = EncColumn(PlainIn)

'initialize 2nd columnar key
keyCol2 = TrimText(keyCol2, True, False, False, False)
If InitColumnar(keyCol2) <> 0 Then Exit Function

'encode
EncodeDoubleColumnar = EncColumn(EncodeDoubleColumnar)

End Function


Public Function DecodeDoubleColumnar(ByVal CodeIn As String, ByVal keyCol1 As String, ByVal keyCol2 As String) As String
'encode with Double Columnar

'trim all but alphabet
CodeIn = TrimText(CodeIn, True, False, False, False)
If CodeIn = "" Then Exit Function

'initialize 2nd columnar key
keyCol2 = TrimText(keyCol2, True, False, False, False)
If InitColumnar(keyCol2) <> 0 Then Exit Function

'decode
DecodeDoubleColumnar = DecColumn(CodeIn)

'initialize 1st columnar key
keyCol1 = TrimText(keyCol1, True, False, False, False)
If InitColumnar(keyCol1) <> 0 Then Exit Function

'decode
DecodeDoubleColumnar = DecColumn(DecodeDoubleColumnar)

End Function


Public Function InitColumnar(ByVal key As String) As Integer
'initialize the columnar key

Dim i As Long
Dim j As Long
Dim PWL As Integer
Dim smallestChar As Byte
Dim currentChar As Byte

'check Key
PWL = Len(key)
If PWL < 5 Then
    MsgBox "The Columnar Key is too short.", vbCritical
    InitColumnar = 1
    Exit Function
    End If

'Get Key column order and put in in row()
ReDim Row(PWL) As Integer
For i = 1 To PWL
    smallestChar = 255
    For j = 1 To PWL
        currentChar = Asc(UCase(Mid(key, j, 1)))
        If currentChar < smallestChar Then
            smallestChar = currentChar
            Row(i) = j

        End If
    Next
    Mid(key, Row(i), 1) = Chr(255)
Next

End Function


Public Function EncColumn(ByVal PlainIn As String) As String
'encode text columnar

Dim i As Long
Dim j As Long

'readoff row by row and place one by one
For i = 1 To UBound(Row)
    For j = Row(i) To Len(PlainIn) Step UBound(Row)
        EncColumn = EncColumn & Mid(PlainIn, j, 1)
    Next
Next

End Function


Public Function DecColumn(ByVal CodeIn As String) As String
'decode text columnar

Dim i As Long
Dim j As Long
Dim CodeCount As Long

'readoff one by one and place row by row
DecColumn = Space(Len(CodeIn))
CodeCount = 1
For i = 1 To UBound(Row)
    For j = Row(i) To Len(CodeIn) Step UBound(Row)
        Mid(DecColumn, j, 1) = Mid(CodeIn, CodeCount, 1)
        CodeCount = CodeCount + 1
    Next
Next

End Function

'----------------------------------------------------------------
'
'              Straddling Checkerboard Subs
'
'----------------------------------------------------------------

Public Function EncodeCheckerBoard(ByVal PlainIn As String, ByVal key As String) As String
'encode with checkerboard

PlainIn = TrimText(PlainIn, True, False, True, True)
If PlainIn = "" Then Exit Function

'initialize CheckerBoard key
key = TrimText(key, True, False, False, False)
If InitCheckerboard(key) <> 0 Then Exit Function

'encode
EncodeCheckerBoard = EncChecker(PlainIn)

End Function


Public Function DecodeCheckerBoard(ByVal CodeIn As String, ByVal key As String) As String
'decode with checkerboard

'trim all but alphabet
CodeIn = TrimText(CodeIn, False, True, False, False)
If CodeIn = "" Then Exit Function

'initialize CheckerBoard key
key = TrimText(key, True, False, False, False)
If InitCheckerboard(key) <> 0 Then Exit Function

'decode
DecodeCheckerBoard = DecChecker(CodeIn)

End Function


Private Function InitCheckerboard(ByVal key As String) As Integer
'initialize checkerboard key

Dim i As Long
Dim j As Long
Dim smallestChar As Byte
Dim currentChar As Byte
Dim smallestPointer As Integer
Dim LO As Byte
Dim HI As Byte
Dim Row(10) As Integer

'check key and text lenght
If Len(key) < 10 Then
    MsgBox "The Checkerboard Key must be at least 10 characters.", vbCritical
    InitCheckerboard = 1
    Exit Function
    End If
    
' assign codes to standard numbered checkerboard
Plain = Array("", "1", "31", "32", "33", "6", "34", "35", _
        "36", "9", "37", "38", "39", "30", "5", "4", "71", _
        "72", "0", "8", "2", "73", "74", "75", "76", "77", _
        "78", "79", "70")

'Get Key column order
For i = 1 To 10
    smallestChar = 255
    For j = 1 To 10
        currentChar = Asc(UCase(Mid(key, j, 1)))
        If currentChar < smallestChar Then
            smallestChar = currentChar
            smallestPointer = j
        End If
    Next
    numConv(smallestPointer Mod 10) = i Mod 10
    Mid(key, smallestPointer, 1) = Chr(255)
Next

'setup re-ordered checkerboard numbers
For i = 1 To 28
    If Len(Plain(i)) = 1 Then
        LO = Val(Plain(i))
        Code(i) = Trim(Str(numConv(LO)))
        Else
        LO = Val(Right(Plain(i), 1))
        HI = Val(Left(Plain(i), 1))
        Code(i) = Trim(Str(numConv(HI))) & Trim(Str(numConv(LO)))
        End If
Next i

End Function


Public Function EncChecker(ByVal PlainIn As String) As String
'encode text checkerboard
Dim i As Long

For i = 1 To Len(PlainIn)
    EncChecker = EncChecker & GetCode(Mid(PlainIn, i, 1))
Next i

End Function

Public Function DecChecker(ByVal CodeIn As String) As String
'decode text checkerboard
Dim i As Long
Dim Pchar As String

For i = 1 To Len(CodeIn)
    Pchar = GetPlain(Mid(CodeIn, i, 1))
    If Pchar = "" Then
        Pchar = GetPlain(Mid(CodeIn, i, 2))
        i = i + 1
    End If
    DecChecker = DecChecker & Pchar
Next

End Function


Private Function GetCode(PlainChar As String) As String
'find number that matches to character

Dim x As Byte
x = Asc(UCase(PlainChar))
If x = Asc(".") Then
    GetCode = Code(27) ' point
ElseIf x = Asc(" ") Then
    GetCode = Code(28) ' space
ElseIf x > 64 And x < 91 Then
    GetCode = Code(x - 64) ' letter
Else
    GetCode = "" ' not found
End If

End Function


Private Function GetPlain(CodeChar As String) As String
'find character that matches to number

Dim i As Integer

For i = 1 To 28
    If CodeChar = Code(i) Then
        'match found
        If i = 27 Then
            GetPlain = "." ' point
        ElseIf i = 28 Then
            GetPlain = " " ' space
        Else
            GetPlain = Chr(i + 64) ' letter
        End If
    Exit Function
    End If
Next

GetPlain = ""

End Function


Public Function EncodeCheckAndColumnar(ByVal PlainIn As String, ByVal KeySCB As String, ByVal keyCol1 As String, ByVal keyCol2 As String) As String
'Encode CheckerBoard with Double Columnar

PlainIn = TrimText(PlainIn, True, False, True, True)
If PlainIn = "" Then Exit Function

'initialize checkerboard key
KeySCB = TrimText(KeySCB, True, False, False, False)
If InitCheckerboard(KeySCB) <> 0 Then Exit Function

'encode
EncodeCheckAndColumnar = EncChecker(PlainIn)
If EncodeCheckAndColumnar = "" Then Exit Function

'initialize 1st columnar key
keyCol1 = TrimText(keyCol1, True, False, False, False)
If InitColumnar(keyCol1) <> 0 Then Exit Function

'encode
EncodeCheckAndColumnar = EncColumn(EncodeCheckAndColumnar)
If EncodeCheckAndColumnar = "" Then Exit Function

'initialize 2nd columnar key
keyCol2 = TrimText(keyCol2, True, False, False, False)
If InitColumnar(keyCol2) <> 0 Then Exit Function

'encode
EncodeCheckAndColumnar = EncColumn(EncodeCheckAndColumnar)

End Function


Public Function DecodeCheckAndColumnar(ByVal CodeIn As String, ByVal KeySCB As String, ByVal keyCol1 As String, ByVal keyCol2 As String) As String
'decode CheckerBoard with Double Columnar

'trim all but alphabet
CodeIn = TrimText(CodeIn, False, True, False, False)
If CodeIn = "" Then Exit Function

'initialize 2nd columnar key
keyCol2 = TrimText(keyCol2, True, False, False, False)
If InitColumnar(keyCol2) <> 0 Then Exit Function

'decode
DecodeCheckAndColumnar = DecColumn(CodeIn)
If DecodeCheckAndColumnar = "" Then Exit Function

'initialize 1st columnar key
keyCol1 = TrimText(keyCol1, True, False, False, False)
If InitColumnar(keyCol1) <> 0 Then Exit Function

'decode
DecodeCheckAndColumnar = DecColumn(DecodeCheckAndColumnar)

'initialize checkerboard key
KeySCB = TrimText(KeySCB, True, False, False, False)
If InitCheckerboard(KeySCB) <> 0 Then Exit Function

'decode
DecodeCheckAndColumnar = DecChecker(DecodeCheckAndColumnar)

End Function


'----------------------------------------------------------------
'
'                           ADFGVX Subs
'
'----------------------------------------------------------------

Public Function EncodeADFGVX(ByVal PlainIn As String, ByVal KeySquare As String, ByVal KeyCol As String) As String
'Encode with ADFGVX

PlainIn = TrimText(PlainIn, True, True, False, False)
If PlainIn = "" Then Exit Function

'initialize Square key
KeySquare = TrimText(KeySquare, True, False, False, False)
If InitSquare(KeySquare) <> 0 Then Exit Function

'encode
EncodeADFGVX = EncSquare(PlainIn)
If EncodeADFGVX = "" Then Exit Function

'initialize columnar key
KeyCol = TrimText(KeyCol, True, False, False, False)
If InitColumnar(KeyCol) <> 0 Then Exit Function

'encode
EncodeADFGVX = EncColumn(EncodeADFGVX)

End Function


Public Function DecodeADFGVX(ByVal CodeIn As String, ByVal KeySquare As String, ByVal KeyCol As String) As String
'Decode with ADFGVX

'trim all but alphabet
CodeIn = TrimText(CodeIn, True, False, False, False)
If CodeIn = "" Then Exit Function

'initialize columnar key
KeyCol = TrimText(KeyCol, True, False, False, False)
If InitColumnar(KeyCol) <> 0 Then Exit Function

'decode column
DecodeADFGVX = DecColumn(CodeIn)
If DecodeADFGVX = "" Then Exit Function

'initialize square key
KeySquare = TrimText(KeySquare, True, False, False, False)
If InitSquare(KeySquare) <> 0 Then Exit Function

'decode square
DecodeADFGVX = DecSquare(DecodeADFGVX)

End Function

Private Function InitSquare(key As String) As Integer
'initialize ADFGVX key

Dim i As Integer
Dim SquareKey As String
Dim SQ As String
Dim SquarePos As Integer

'check key and text lenght
If Len(key) < 3 Then
    MsgBox "The Square Key is too small.", vbCritical
    InitSquare = 1
    Exit Function
    End If

'delete doubles in key
SquareKey = Left(key, 1)
For i = 2 To Len(key)
    SQ = Mid(key, i, 1)
    If InStr(1, SquareKey, SQ) = 0 Then SquareKey = SquareKey & SQ
Next

'fill rest of key
For i = 1 To 26
    SQ = Chr(i + 64)
    If InStr(1, SquareKey, SQ) = 0 Then SquareKey = SquareKey & SQ
Next

'fill key and figures in square
SquarePos = 1
For i = 1 To 26
    SQ = Mid(SquareKey, i, 1)
    Square(SquarePos) = SQ
    If Asc(SQ) > 64 And Asc(SQ) < 75 Then
        'after letter comes number
        SquarePos = SquarePos + 1
        If Asc(SQ) = 74 Then
            'after J comes zero
            Square(SquarePos) = Chr(Asc(SQ) + 30)
            Else
            'after A comes 1, after B comes 2 etc...
            Square(SquarePos) = Chr(Asc(SQ) - 16)
            End If
        Else
        Square(SquarePos) = SQ
        End If
    SquarePos = SquarePos + 1
Next

'set column and row headers
SquareCode(0) = "A"
SquareCode(1) = "D"
SquareCode(2) = "F"
SquareCode(3) = "G"
SquareCode(4) = "V"
SquareCode(5) = "X"

End Function


Private Function EncSquare(PlainIn As String) As String
'encode ADFGVX square

Dim i As Integer
Dim j As Integer
Dim x As Integer
Dim Y As Integer

For i = 1 To Len(PlainIn)
    For j = 1 To 36
        'search for matching letter or number in key square
        If Mid(PlainIn, i, 1) = Square(j) Then
            'get row and column
            Y = Int((j - 1) / 6)
            x = (j - 1) - (Y * 6)
            'encode to ADFGVX letter
            EncSquare = EncSquare & SquareCode(Y) & SquareCode(x)
        End If
    Next
Next

End Function


Private Function DecSquare(CodeIn As String) As String
'decode ADFGVX square

Dim i As Integer
Dim x As Integer
Dim Y As Integer

'read off in groups of two (XY)
For i = 1 To Len(CodeIn) Step 2
    'get row and column of ADFGVX letter
    Y = GetADFGVXcode(Mid(CodeIn, i, 1))
    x = GetADFGVXcode(Mid(CodeIn, i + 1, 1))
    'get the decode letter in the key square
    DecSquare = DecSquare & Square((Y * 6) + x + 1)
Next

End Function


Private Function GetADFGVXcode(CharIn As String) As Integer
'get the number value of one of the ADFGVX letters
Dim i As Integer

For i = 0 To 5
    If CharIn = SquareCode(i) Then GetADFGVXcode = i
Next i

End Function

'----------------------------------------------------------------
'
'                       Vigenére Subs
'
'----------------------------------------------------------------


Public Function EncodeVigenere(ByVal PlainIn As String, ByVal key As String) As String
'Encode with vigenere

Dim i As Long
Dim Cin As Integer
Dim Ckey As Integer
Dim Cout As Integer
Dim Keypos As Integer

key = TrimText(key, True, False, False, False)
If Len(key) < 2 Then
    MsgBox "Key size too small", vbCritical
    Exit Function
    End If

PlainIn = TrimText(PlainIn, True, False, False, False)
If PlainIn = "" Then Exit Function

Keypos = 1
For i = 1 To Len(PlainIn)
    Cin = Asc(Mid(PlainIn, i, 1)) - 64
    Ckey = Asc(Mid(key, Keypos, 1)) - 64
    Cout = Cin + (Ckey - 1)
    If Cout > 26 Then Cout = Cout - 26
    EncodeVigenere = EncodeVigenere & Chr(Cout + 64)
    Keypos = Keypos + 1: If Keypos > Len(key) Then Keypos = 1
Next i

End Function


Public Function DecodeVigenere(ByVal PlainIn As String, ByVal key As String)
'Encode with vigenere

Dim i As Long
Dim Cin As Integer
Dim Ckey As Integer
Dim Cout As Integer
Dim Keypos As Integer

key = TrimText(key, True, False, False, False)
If Len(key) < 2 Then
    MsgBox "Key size too small", vbCritical
    Exit Function
    End If

PlainIn = TrimText(PlainIn, True, False, False, False)
If PlainIn = "" Then Exit Function

Keypos = 1
For i = 1 To Len(PlainIn)
    Cin = Asc(Mid(PlainIn, i, 1)) - 64
    Ckey = Asc(Mid(key, Keypos, 1)) - 64
    Cout = Cin - (Ckey - 1)
    If Cout < 1 Then Cout = Cout + 26
    DecodeVigenere = DecodeVigenere & Chr(Cout + 64)
    Keypos = Keypos + 1: If Keypos > Len(key) Then Keypos = 1
Next i

End Function

'----------------------------------------------------------------
'
'                       Playfair Subs
'
'----------------------------------------------------------------


Public Function EncodePlayFair(ByVal PlainIn As String, ByVal key As String)
'encode with plaifair

Dim i As Long
Dim P1 As String
Dim P2 As String
Dim Bpos As Long
Dim tmpText As String

PlainIn = TrimText(PlainIn, True, False, False, False)
If PlainIn = "" Then Exit Function

Bpos = 1
Do
    'replace J's by I's
    If Mid(PlainIn, Bpos, 1) = "J" Then Mid(PlainIn, Bpos, 1) = "I"
    If Mid(PlainIn, Bpos + 1, 1) = "J" Then Mid(PlainIn, Bpos + 1, 1) = "I"
    'check for double-letter bigrams
    If Mid(PlainIn, Bpos, 1) <> Mid(PlainIn, Bpos + 1, 1) Then
        'bigram ok
        Bpos = Bpos + 2
        Else
        'bigram two identical letters, so insert X
        PlainIn = Left(PlainIn, Bpos) & "X" & Mid(PlainIn, Bpos + 1)
        Bpos = Bpos + 2
    End If
Loop While Bpos < Len(PlainIn)

'make even textlenght
If Len(PlainIn) Mod 2 <> 0 Then PlainIn = PlainIn & "X"

'initialize key
key = TrimText(key, True, False, False, False)
If Len(key) < 2 Then
    MsgBox "Key size too small", vbCritical
    Exit Function
    End If
If InitPlayFair(key) <> 0 Then Exit Function

For i = 1 To Len(PlainIn) Step 2
    P1 = Mid(PlainIn, i, 1)
    P2 = Mid(PlainIn, i + 1, 1)
    EncodePlayFair = EncodePlayFair & EncodeDigram(P1, P2)
Next

End Function


Public Function DecodePlayFair(ByVal CodeIn As String, ByVal key As String)
'decode with plaifair

Dim i As Long
Dim P1 As String
Dim P2 As String

CodeIn = TrimText(CodeIn, True, False, False, False)
If CodeIn = "" Then Exit Function

'initialize key
key = TrimText(key, True, False, False, False)
If Len(key) < 2 Then
    MsgBox "Key size too small", vbCritical
    Exit Function
    End If
If InitPlayFair(key) <> 0 Then Exit Function

If Len(CodeIn) Mod 2 <> 0 Then
    MsgBox "Impossible to split text into Digrams", vbCritical
    Exit Function
    End If
    
For i = 1 To Len(CodeIn) Step 2
    P1 = Mid(CodeIn, i, 1)
    P2 = Mid(CodeIn, i + 1, 1)
    DecodePlayFair = DecodePlayFair & DecodeDigram(P1, P2)
Next

End Function


Private Function EncodeDigram(ByVal P1 As String, ByVal P2 As String) As String
Dim X1 As Integer
Dim Y1 As Integer
Dim X2 As Integer
Dim Y2 As Integer
Dim tmpX As Integer
Dim tmpY As Integer

Call GetXY(P1, X1, Y1)
Call GetXY(P2, X2, Y2)

If X1 = X2 Then
    'same column
    Y1 = Y1 + 1: If Y1 > 4 Then Y1 = Y1 - 5
    Y2 = Y2 + 1: If Y2 > 4 Then Y2 = Y2 - 5
ElseIf Y1 = Y2 Then
    'same row
    X1 = X1 + 1: If X1 > 4 Then X1 = X1 - 5
    X2 = X2 + 1: If X2 > 4 Then X2 = X2 - 5
Else
    'different col and row (Z methode)
    tmpX = X1
    tmpY = Y1
    X1 = X2
    X2 = tmpX
End If

P1 = GetXYchar(X1, Y1)
P2 = GetXYchar(X2, Y2)

EncodeDigram = P1 & P2

End Function


Private Function DecodeDigram(ByVal P1 As String, ByVal P2 As String) As String
Dim X1 As Integer
Dim Y1 As Integer
Dim X2 As Integer
Dim Y2 As Integer
Dim tmpX As Integer
Dim tmpY As Integer

Call GetXY(P1, X1, Y1)
Call GetXY(P2, X2, Y2)

If X1 = X2 Then
    'same column
    Y1 = Y1 - 1: If Y1 < 0 Then Y1 = Y1 + 5
    Y2 = Y2 - 1: If Y2 < 0 Then Y2 = Y2 + 5
ElseIf Y1 = Y2 Then
    'same row
    X1 = X1 - 1: If X1 < 0 Then X1 = X1 + 5
    X2 = X2 - 1: If X2 < 0 Then X2 = X2 + 5
Else
    'different col and row (Z methode)
    tmpX = X1
    tmpY = Y1
    X1 = X2
    X2 = tmpX
End If

P1 = GetXYchar(X1, Y1)
P2 = GetXYchar(X2, Y2)

DecodeDigram = P1 & P2

End Function


Private Function GetXY(Pchar As String, x As Integer, Y As Integer)
'find X and Y from a character
Dim PosP As Integer

PosP = InStr(1, PlaySquare, Pchar) - 1
Y = Int(PosP / 5)
x = PosP - (Y * 5)

End Function


Private Function GetXYchar(x As Integer, Y As Integer)
'get the char by X and Y
GetXYchar = Mid(PlaySquare, (Y * 5) + x + 1, 1)
End Function


Public Function InitPlayFair(ByVal key As String) As Integer
Dim i As Integer
Dim SQ As String

PlaySquare = ""
'delete doubles in key
For i = 1 To Len(key)
    SQ = Mid(key, i, 1)
    If InStr(1, PlaySquare, SQ) = 0 And SQ <> "J" Then PlaySquare = PlaySquare & SQ
Next

'fill rest of key
For i = 1 To 26
    SQ = Chr(i + 64)
    If InStr(1, PlaySquare, SQ) = 0 And SQ <> "J" Then PlaySquare = PlaySquare & SQ
Next

End Function

'----------------------------------------------------------------
'
'                           General Subs
'
'----------------------------------------------------------------

Public Function TrimText(TextIn As String, Letters As Boolean, Numbers As Boolean, Spaces As Boolean, Points As Boolean)
'trim a strings letters, numbers, spaces or points
Dim i As Long
Dim tmp As Byte
For i = 1 To Len(TextIn)
    tmp = Asc(UCase(Mid(TextIn, i, 1)))
    If Letters = True And (tmp > 64 And tmp < 123) Then
        TrimText = TrimText & Chr(tmp)
    ElseIf Numbers = True And (tmp > 47 And tmp < 58) Then
        TrimText = TrimText & Chr(tmp)
    ElseIf Spaces = True And tmp = 32 Then
        TrimText = TrimText & Chr(tmp)
    ElseIf Points = True And tmp = 46 Then
        TrimText = TrimText & Chr(tmp)
    End If
Next
End Function

Public Function MakeGroups(TextIn As String, Groups As Boolean, GroupsPerLine As Integer) As String
'devide code text in groups
Dim i As Long
If Groups = False Or GroupsPerLine = 0 Then MakeGroups = TextIn: Exit Function
For i = 1 To Len(TextIn)
    MakeGroups = MakeGroups & Mid(TextIn, i, 1)
    If i Mod 4 = 0 And i <> Len(TextIn) Then MakeGroups = MakeGroups & "-"
    If i Mod (GroupsPerLine * 4) = 0 Then MakeGroups = MakeGroups & vbCrLf
Next
End Function

Public Function TestKey(aKey As Integer, keyNumbers As Boolean, keySpace As Boolean, keyPoint As Boolean) As Integer
'returns only allowed characters
If aKey > 64 And aKey < 91 Then
    TestKey = aKey
ElseIf aKey > 96 And aKey < 123 Then
    TestKey = aKey - 32
ElseIf (aKey > 47 And aKey < 58) And keyNumbers = True Then
    TestKey = aKey
ElseIf aKey = 32 And keySpace = True Then
    TestKey = aKey
ElseIf aKey = 46 And keyPoint = True Then
    TestKey = aKey
ElseIf aKey < 32 Then
    TestKey = aKey
Else
TestKey = 0
End If
End Function







