VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrickMP3Player"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type WNDCLASSEX
    cbSize              As Long
    style               As Long
    lpfnwndproc         As Long
    cbClsextra          As Long
    cbWndExtra2         As Long
    hInstance           As Long
    hIcon               As Long
    hCursor             As Long
    hbrBackground       As Long
    lpszMenuName        As Long
    lpszClassName       As Long
    hIconSm             As Long
End Type

Private Type MPEGLAYER3WAVEFORMAT
    wFormatTag          As Integer
    nChannels           As Integer
    nSamplesPerSec      As Long
    nAvgBytesPerSec     As Long
    nBlockAlign         As Integer
    wBitsPerSample      As Integer
    cbSize              As Integer
    wID                 As Integer
    fdwFlags            As Long
    nBlockSize          As Integer
    nFramesPerBlock     As Integer
    nCodecDelay         As Integer
End Type

Private Type FrameInfo
    offset              As Long
    bitrate             As Long
End Type

Private Type Mp3Info
    format              As MPEGLAYER3WAVEFORMAT
    lpFrameOffset       As Long
    szDataSize          As Long
    samplesPerFrame     As Long
    framesCount         As Long
    frameOffset()       As FrameInfo
End Type

Private Type WAVEHDR
    lpData              As Long
    dwBufferLength      As Long
    dwBytesRecorded     As Long
    dwUser              As Long
    dwFlags             As Long
    dwLoops             As Long
    lpNext              As Long
    Reserved            As Long
End Type
 
Private Type mp3Buffer
    header              As WAVEHDR
    status              As Boolean
End Type

Private Type mp3Const
    bitrate(1, 15)      As Integer
    smprate(2, 3)       As Long
End Type

Private Type curBuffer
    b(15)               As Currency
End Type

Private Declare Function GetMem4 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare Function GetMem8 Lib "msvbvm60" (Src As Any, Dst As Any) As Long
Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleW" (ByVal lpModuleName As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapCreate Lib "kernel32" (ByVal flOptions As Long, ByVal dwInitialSize As Long, ByVal dwMaximumSize As Long) As Long
Private Declare Function HeapDestroy Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function SetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableW" (ByVal lpName As Long, lpValue As Any) As Long
Private Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableW" (ByVal lpName As Long, lpBuffer As Any, ByVal nSize As Long) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassW" (ByVal lpClassName As Long, ByVal hInstance As Long) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExW" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowSubclass Lib "Comctl32" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, dwRefData As Any) As Long
Private Declare Function RemoveWindowSubclass Lib "Comctl32" (ByVal hwnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function DefSubclassProc Lib "Comctl32" (ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function waveOutOpen Lib "winmm" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As Any, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function waveOutPrepareHeader Lib "winmm" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutWrite Lib "winmm" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutUnprepareHeader Lib "winmm" (ByVal hWaveOut As Long, lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function waveOutClose Lib "winmm" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutReset Lib "winmm" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutPause Lib "winmm" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutRestart Lib "winmm" (ByVal hWaveOut As Long) As Long
Private Declare Function waveOutSetVolume Lib "winmm" (ByVal wDeviceID As Long, ByVal dwVolume As Long) As Long
Private Declare Function waveOutGetVolume Lib "winmm" (ByVal wDeviceID As Long, dwVolume As Long) As Long

Private Const Mp3Class                      As String = "TrickMP3PlayerClass"
Private Const HWND_MESSAGE                  As Long = -3
Private Const WAVE_MAPPER                   As Long = -1&
Private Const WHDR_DONE                     As Long = &H1
Private Const CALLBACK_WINDOW               As Long = &H10000
Private Const MM_WOM_DONE                   As Long = &H3BD
Private Const WM_TIMER                      As Long = &H113
Private Const WNDPROCINDEX                  As Long = 13
Private Const HEAP_CREATE_ENABLE_EXECUTE    As Long = &H40000
Private Const HEAP_NO_SERIALIZE             As Long = &H1
Private Const HEAP_ZERO_MEMORY              As Long = &H8
Private Const GWL_USERDATA                  As Long = (-21)
Private Const MPEGLAYER3_FLAG_PADDING_OFF   As Long = 2
Private Const WAVE_FORMAT_MPEGLAYER3        As Long = &H55
Private Const MPEGLAYER3_WFX_EXTRA_BYTES    As Long = 12
Private Const MPEGLAYER3_ID_MPEG            As Long = 1
Private Const BUFFERS_COUNT                 As Long = 8

Private init        As Boolean
Private loaded      As Boolean
Private playing     As Boolean
Private paused      As Boolean
Private isLoop      As Boolean
Private constants   As mp3Const
Private hwnd        As Long
Private hHeap       As Long
Private lpWndProc   As Long
Private hWave       As Long
Private headers()   As mp3Buffer
Private curPosition As Long
Private fileInfo    As Mp3Info
Private buffer()    As Byte
Private mPan        As Single
Private mVolume     As Single

' // Initialize playback. The first parameter is a pointer to data of the raw mp3 file.
' // Second parameter is a size of this file in bytes.
' // Last parameter indicates that need to copy this file in the internal buffer.
Public Function Initialize(ByVal lpData As Long, ByVal szData As Long, Optional ByVal blCopy As Boolean) As Boolean
    Dim status  As Boolean
    Dim info    As Mp3Info
    Dim ret     As Long
    Dim index   As Long
    
    If Not init Then Exit Function
    
    status = Mp3GetInfo(lpData, szData, info)
    If Not status Then Exit Function
    
    If hWave Then ClearAll
    
    If blCopy Then
        
        ReDim buffer(info.szDataSize - 1)
        memcpy buffer(0), ByVal info.lpFrameOffset, info.szDataSize
        info.lpFrameOffset = VarPtr(buffer(0))
        
    End If
    
    ret = waveOutOpen(hWave, WAVE_MAPPER, info.format, hwnd, 0, CALLBACK_WINDOW)
    If ret Then hWave = 0:  Exit Function

    fileInfo = info
    curPosition = 0
    Me.Pan = mPan
    Me.Volume = mVolume
    
    loaded = True
    playing = False
    
End Function

' // Start playback. If it is the first call after stopping or initialization then parameter "looped" allows to play a data by circularly.
Public Function Play(Optional ByVal looped As Boolean) As Boolean
    Dim index   As Long
    Dim ret     As Long
    
    If Not (init And loaded) Then Exit Function
    
    isLoop = looped
    
    If paused Then
        
        If waveOutRestart(hWave) Then Exit Function
        paused = False
        
    Else
        
        curPosition = 0
        
        For index = 0 To BUFFERS_COUNT - 1
    
            headers(index).header.lpData = fileInfo.lpFrameOffset + fileInfo.frameOffset(curPosition).offset
    
            If index < fileInfo.framesCount - 1 Then
            
                headers(index).header.dwBufferLength = fileInfo.frameOffset(curPosition + 1).offset - fileInfo.frameOffset(curPosition).offset
                
            Else

                headers(index).header.dwBufferLength = fileInfo.szDataSize - fileInfo.frameOffset(curPosition).offset
                
                If isLoop Then
                    curPosition = 0
                Else
                    Exit For
                End If
                
            End If
    
            ret = waveOutPrepareHeader(hWave, headers(index).header, Len(headers(index).header))
            headers(index).status = ret = 0
    
            If ret Then ClearAll: Exit Function
    
            ret = waveOutWrite(hWave, headers(index).header, Len(headers(index).header))
            If ret Then ClearAll: Exit Function
            
            curPosition = curPosition + 1
            
        Next
        
    End If
    
    playing = True
    Play = True
    
End Function

' // Pause playback.
Public Function Pause() As Boolean

    If Not (init And loaded And playing) Then Exit Function
    
    waveOutPause hWave
    
    paused = True
    Pause = True
    
End Function

' // Stop playback.
Public Function StopPlaying() As Boolean

    If Not (init And loaded And playing) Then Exit Function
    
    paused = False
    playing = False
    curPosition = -1
    
    waveOutReset hWave

    StopPlaying = True
    
End Function

' // Set current playback position (in milliseconds).
Public Function SetPositionMs(ByVal pos As Long) As Boolean
    Dim frameLength As Single
    Dim index       As Long
    
    If Not (init And loaded) Then Err.Raise 5: Exit Function
    
    frameLength = fileInfo.samplesPerFrame / fileInfo.format.nSamplesPerSec
    index = pos / 1000 / frameLength
    
    If index >= fileInfo.framesCount Then Err.Raise 5:  Exit Function
    
    curPosition = index
    SetPositionMs = True
    
End Function

' // Get current playback position (in milliseconds).
Public Function GetPositionMs() As Long
    Dim frameLength As Single
    
    If Not (init And loaded) Then Exit Function
    
    frameLength = fileInfo.samplesPerFrame / fileInfo.format.nSamplesPerSec
    GetPositionMs = curPosition * frameLength * 1000
    
End Function

' // Get duration of the data in milliseconds.
Public Function GetDurationMs() As Long
    Dim frameLength As Single
    
    If Not (init And loaded) Then Exit Function
    
    frameLength = fileInfo.samplesPerFrame / fileInfo.format.nSamplesPerSec
    GetDurationMs = fileInfo.framesCount * frameLength * 1000
    
End Function

' // Get current bitrate.
Public Function GetBitrate() As Long
    
    If curPosition < 0 Then Exit Function
    GetBitrate = fileInfo.frameOffset(curPosition).bitrate
    
End Function

' // If playback is active then true.
Public Property Get IsPlaying() As Boolean
    IsPlaying = init And loaded And playing And Not paused
End Property

' // Volume
Public Property Get Volume() As Single
    Dim dwVolume    As Long
    Dim volLeft     As Long
    Dim volRight    As Long
    
    waveOutGetVolume hWave, dwVolume
    
    volLeft = dwVolume And &HFFFF&
    volRight = ((dwVolume And &HFFFF0000) \ &H10000) And &HFFFF&
    
    If volLeft > volRight Then Volume = volLeft / 65535 Else Volume = volRight / 65535
    
End Property
Public Property Let Volume(ByVal value As Single)
    Dim dwVolume    As Long
    Dim volRight    As Long
    
    If value > 1 Or value <= 0 Then Err.Raise 6: Exit Property
    
    mVolume = value
    
    If mPan > 0 Then
        volRight = value * 65535
        dwVolume = volRight * (1 - mPan)
    Else
        dwVolume = value * 65535
        volRight = dwVolume * (1 + mPan)
    End If
        
    If volRight And &H8000& Then
        dwVolume = dwVolume Or ((volRight And &H7FFF&) * &H10000) Or &H80000000
    Else
        dwVolume = dwVolume Or (volRight * &H10000)
    End If
    
    waveOutSetVolume hWave, dwVolume
     
End Property

' // Pan
Public Property Get Pan() As Single
    Dim dwVolume    As Long
    Dim volLeft     As Long
    Dim volRight    As Long
    
    waveOutGetVolume hWave, dwVolume
    
    volLeft = dwVolume And &HFFFF&
    volRight = ((dwVolume And &HFFFF0000) \ &H10000) And &HFFFF&
    If volLeft > volRight Then dwVolume = volLeft Else dwVolume = volRight
    
    If dwVolume = 0 Then dwVolume = 1
    Pan = (volRight - volLeft) / dwVolume
    
End Property
Public Property Let Pan(ByVal value As Single)
    
    If value > 1 Or value < -1 Then Err.Raise 6: Exit Property
    
    mPan = value
    Me.Volume = mVolume
    
End Property

' // Local procedures.
Private Function SUBCLASSPROC(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Dim index   As Long
    Dim lpData  As Long
    Dim inIDE   As Boolean
    
    Debug.Assert MakeTrue(inIDE)
    
    If inIDE Then
    
        If Msg = WM_TIMER Then
        
            KillTimer hwnd, wParam
            
            For index = 0 To BUFFERS_COUNT - 1
                
                If headers(index).header.dwFlags And WHDR_DONE Then
                    
                    WriteNext index
                    
                End If
                
            Next
            
        End If
        
    End If
    
    If Msg = MM_WOM_DONE Then
        
        If wParam <> hWave Then GoTo DefCall
        
        GetMem4 ByVal lParam, lpData

        index = GetBufferIndex(lpData)

        If index = -1 Then GoTo DefCall
        
        WriteNext index
        
    End If
    
DefCall:
    
    SUBCLASSPROC = DefSubclassProc(hwnd, Msg, wParam, lParam)
    
End Function

Private Sub WriteNext(ByVal index As Long)
    
    waveOutUnprepareHeader hWave, headers(index).header, Len(headers(index).header)
    
    If playing = False And paused = False Then Exit Sub
    
    If curPosition = -1 Then Exit Sub
    
    headers(index).header.dwFlags = headers(index).header.dwFlags And Not WHDR_DONE
    headers(index).header.lpData = fileInfo.lpFrameOffset + fileInfo.frameOffset(curPosition).offset

    If curPosition < fileInfo.framesCount - 1 Then
    
        headers(index).header.dwBufferLength = fileInfo.frameOffset(curPosition + 1).offset - fileInfo.frameOffset(curPosition).offset
        curPosition = curPosition + 1
        
    Else
    
        headers(index).header.dwBufferLength = fileInfo.szDataSize - fileInfo.frameOffset(curPosition).offset
        
        If isLoop Then
            curPosition = 0
        Else
            curPosition = -1
        End If
            
    End If
    
    waveOutPrepareHeader hWave, headers(index).header, Len(headers(index).header)
    waveOutWrite hWave, headers(index).header, Len(headers(index).header)
    
End Sub

Private Sub ClearAll()
    Dim index   As Long

    If hWave = 0 Then Exit Sub

    For index = 0 To BUFFERS_COUNT - 1

        If headers(index).status Then
            waveOutUnprepareHeader hWave, headers(index).header, Len(headers(index).header)
        End If

    Next
    
    If playing Or paused Then waveOutReset hWave
    
    waveOutClose hWave
    
    loaded = False
    playing = False
    paused = False
    hWave = 0
    
End Sub

Private Function GetBufferIndex(ByVal ptr As Long) As Long
    Dim index As Long

    For index = 0 To UBound(headers)

        If headers(index).header.lpData = ptr Then
            GetBufferIndex = index
            Exit Function
        End If

    Next

    GetBufferIndex = -1
End Function

Private Function Mp3GetInfo(ByVal lpData As Long, ByVal szData As Long, info As Mp3Info) As Boolean
    Dim hdr(9)  As Byte
    Dim size    As Long
    
    If szData >= 128 Then
        ' Skip ID3V1 tag
        memcpy hdr(0), ByVal lpData + szData - 128, 3
        
        If hdr(0) = &H54 And hdr(1) = &H41 And hdr(2) = &H47 Then
            
            szData = szData - 128
            
        End If
        
    End If
    
    ' Skip ID3V2 tags from beginning
    memcpy hdr(0), ByVal lpData, 10
    
    If hdr(0) = &H49 And hdr(1) = &H44 And hdr(2) = &H33 Then
        
        ' footer present
        If hdr(5) And &H10 Then
            szData = szData - 10
        End If
    
        size = hdr(6) * &H200000
        size = size Or (hdr(7) * &H4000&)
        size = size Or (hdr(8) * &H80&)
        size = size Or hdr(9)
        size = size + 10
        
        lpData = lpData + size
        szData = szData - size
            
    Else
        ' Skip ID3V2 tags from end
        memcpy hdr(0), ByVal lpData + szData - 10, 10
        
        If hdr(2) = &H49 And hdr(1) = &H44 And hdr(0) = &H33 Then
            
            szData = szData - 10
            
            size = hdr(6) * &H200000
            size = size Or (hdr(7) * &H4000&)
            size = size Or (hdr(8) * &H80&)
            size = size Or hdr(9)
            size = size + 10
        
            szData = szData - size
            
        End If
        
    End If
    
    If szData < 4 Then Exit Function
    
    info.framesCount = 0
    'Scan headers
    Do
        ' Find a frame sync
        Do
        
            GetMem4 ByVal lpData, hdr(0)
            
            If hdr(0) = &HFF And (hdr(1) And &HE0) = &HE0 Then
                Dim vers    As Long
                Dim layer   As Long
                Dim bitrate As Long
                Dim smprate As Long
                Dim padding As Long
                Dim channel As Long
                               
                vers = (hdr(1) And &H18) \ 8
                If vers = 1 Then Exit Function
    
                layer = (hdr(1) And &H6) \ 2
                If layer <> 1 Then Exit Function ' Only Layer 3
    
                If vers = 3 Then
                    bitrate = constants.bitrate(0, (hdr(2) And &HF0) \ &H10)
                Else
                    bitrate = constants.bitrate(1, (hdr(2) And &HF0) \ &H10)
                End If

                If vers = 3 Then
                    smprate = constants.smprate(0, (hdr(2) And &HC) \ &H4)
                ElseIf vers = 2 Then
                    smprate = constants.smprate(1, (hdr(2) And &HC) \ &H4)
                Else
                    smprate = constants.smprate(2, (hdr(2) And &HC) \ &H4)
                End If
                
                padding = (hdr(2) And &H2) \ 2
                channel = -(((hdr(3) And &HC0) \ 64) <> 3) + 1
                
                If vers = 3 Then
                    size = Int(144000 * bitrate / smprate) + padding
                Else
                    size = Int(72000 * bitrate / smprate) + padding
                End If
                
                With info
                    If .framesCount = 0 Then

                        With .format
                            .wFormatTag = WAVE_FORMAT_MPEGLAYER3
                            .cbSize = MPEGLAYER3_WFX_EXTRA_BYTES
                            .nChannels = channel
                            .nAvgBytesPerSec = bitrate * 128
                            .wBitsPerSample = 0
                            .nBlockAlign = 1
                            .nSamplesPerSec = smprate
                            .nFramesPerBlock = 1
                            .nCodecDelay = 0
                            .fdwFlags = MPEGLAYER3_FLAG_PADDING_OFF
                            .wID = MPEGLAYER3_ID_MPEG
                            .nBlockSize = size
                        End With
                                        
                        .lpFrameOffset = lpData
                        .szDataSize = szData
                        
                        If vers = 3 Then
                            .samplesPerFrame = 1152
                        Else
                            .samplesPerFrame = 576
                        End If
                        
                        ReDim .frameOffset(511)
                    
                    Else
                        
                        If UBound(.frameOffset) = info.framesCount Then
                            ReDim Preserve .frameOffset(UBound(.frameOffset) + 512)
                        End If
                        
                    End If
                    
                    .frameOffset(info.framesCount).offset = lpData - .lpFrameOffset
                    .frameOffset(info.framesCount).bitrate = bitrate
                    
                End With
                
                lpData = lpData + size
                szData = szData - size
                
                Exit Do
                
            End If
            
            lpData = lpData + 1
            szData = szData - 1
            
        Loop While szData >= 4
        
        info.framesCount = info.framesCount + 1
        
    Loop While szData >= 4

    Mp3GetInfo = True
    
End Function

Private Function GetWindowAndHeap(l_hwnd As Long, l_hHeap As Long) As Boolean
    Dim i1      As Long
    Dim i2      As Long
    Dim b       As Long
    Dim arr(16) As Integer
    
    If GetEnvironmentVariable(StrPtr(Mp3Class), arr(0), 32) Then
        
        i1 = 0: i2 = 8
        Do
            If arr(i1) <= &H39 Then b = arr(i1) - &H30 Else b = arr(i1) - &H37
            If l_hHeap And &H8000000 Then l_hHeap = ((l_hHeap And &H7FFFFF) * &H10 Or &H80000000) Or b Else l_hHeap = (l_hHeap * &H10) Or b
            If arr(i2) <= &H39 Then b = arr(i2) - &H30 Else b = arr(i2) - &H37
            If l_hwnd And &H8000000 Then l_hwnd = ((l_hwnd And &H7FFFFF) * &H10 Or &H80000000) Or b Else l_hwnd = (l_hwnd * &H10) Or b
            i1 = i1 + 1: i2 = i2 + 1
        Loop While i1 < 8
        
        GetWindowAndHeap = l_hwnd <> 0 And l_hHeap <> 0
        
    End If

End Function

Private Function SaveWindowAndHeap(ByVal l_hwnd As Long, ByVal l_hHeap As Long) As Boolean
    Dim i1      As Long
    Dim i2      As Long
    Dim b       As Long
    Dim arr(16) As Integer
    
    i1 = 7: i2 = 15
    Do
        b = l_hHeap And &HF
        If b < 10 Then arr(i1) = b + &H30 Else arr(i1) = b + &H37
        b = l_hwnd And &HF
        If b < 10 Then arr(i2) = b + &H30 Else arr(i2) = b + &H37
        l_hHeap = (l_hHeap And &HFFFFFFF0) \ &H10
        l_hwnd = (l_hwnd And &HFFFFFFF0) \ &H10
        i1 = i1 - 1: i2 = i2 - 1
    Loop While i1 >= 0

    SaveWindowAndHeap = SetEnvironmentVariable(StrPtr(Mp3Class), arr(0))
    
End Function

Private Sub Class_Initialize()
    Dim cls         As WNDCLASSEX
    Dim b           As curBuffer
    Dim isFirst     As Boolean
    Dim inIDE       As Boolean
    Dim AsmSize     As Long
    Dim lpAsm       As Long
    Dim lpFlag      As Long
    Dim hInstVB6    As Long
    Dim lpEbMode    As Long
    Dim hInstUser32 As Long
    Dim hComctl32   As Long
    Dim lpDefProc   As Long
    Dim lpSetTimer  As Long
    Dim clearFlag   As Long
        
    b.b(0) = 450377142658.6656@:    b.b(1) = 900743977448.248@:     b.b(2) = 1351114248211.6672@
    b.b(3) = 1801487954948.9248@:   b.b(4) = 2702228496423.3344@:   b.b(5) = 3602975909897.8496@
    b.b(6) = 4503737067267.712@:    b.b(7) = 18941235272.0895@:     b.b(8) = 4735201446.045@
    b.b(9) = 10307921515.2@:        b.b(10) = 13743895348.4@:       b.b(11) = 3435973838.4@
        
    memcpy constants.bitrate(0, 1), b.b(0), 96
    
    ReDim headers(BUFFERS_COUNT - 1)
    mVolume = 1
    
    isFirst = Not GetWindowAndHeap(hwnd, hHeap)

    Debug.Assert MakeTrue(inIDE)
    
    hInstUser32 = GetModuleHandle(StrPtr("user32"))
    
    If inIDE Then
        
        AsmSize = &H65
        
        hInstVB6 = GetModuleHandle(StrPtr("vba6"))
        hComctl32 = GetModuleHandle(StrPtr("hComctl32"))
        If hComctl32 = 0 Then
            hComctl32 = LoadLibrary(StrPtr("Comctl32"))
            If hComctl32 = 0 Then Exit Sub
        End If
        lpEbMode = GetProcAddress(hInstVB6, "EbMode")
        lpDefProc = GetProcAddress(hComctl32, "DefSubclassProc")
        lpSetTimer = GetProcAddress(hInstUser32, "SetTimer")

        b.b(0) = 843073850243758.4259@: b.b(1) = -457424984652572.8729@:    b.b(2) = 2989182470102.0276@
        b.b(3) = -7165957082854.492@:   b.b(4) = -16790531.982@:            b.b(5) = 10059.9531@
        b.b(6) = 116318324260473.7791@: b.b(7) = 116318324260473.7791@:     b.b(8) = 696980420845.4632@
        b.b(9) = 522808547116743.0705@: b.b(10) = 756460495277739.1878@:    b.b(11) = -10565565861.0689@
        b.b(12) = 41538.9951@

        If isFirst Then
            
            lpFlag = HeapAlloc(GetProcessHeap(), HEAP_NO_SERIALIZE Or HEAP_ZERO_MEMORY, 4)
            If lpFlag = 0 Then Exit Sub
        
        Else
            
            lpFlag = GetWindowLong(hwnd, 0)
            
            GetMem4 ByVal lpFlag, clearFlag
            
            If clearFlag Then
                
                DestroyWindow hwnd
                HeapDestroy hHeap
                UnregisterClass StrPtr(Mp3Class), App.hInstance
                
                GetMem4 0&, ByVal lpFlag
                isFirst = True
                
                hwnd = 0
                hHeap = 0
                
                SaveWindowAndHeap 0, 0
                
            End If
            
        End If

    Else
        
        AsmSize = &H20
        
        b.b(0) = 522808547116743.0705@: b.b(1) = 756460495277739.1878@:    b.b(2) = -10565565861.0689@
        b.b(3) = 41538.9951@
        
    End If
    
    If isFirst Then
    
        hHeap = HeapCreate(HEAP_CREATE_ENABLE_EXECUTE Or HEAP_NO_SERIALIZE, 0, 0)
        If hHeap = 0 Then Exit Sub
        
    End If
    
    lpAsm = HeapAlloc(hHeap, HEAP_NO_SERIALIZE, AsmSize)
    
    If lpAsm = 0 Then
        If isFirst Then HeapDestroy hHeap
        Exit Sub
    End If
    
    lpWndProc = lpAsm

    memcpy ByVal lpAsm, b.b(0), AsmSize

    If inIDE Then
    
        GetMem4 lpEbMode - (lpAsm + &H9) - 5, ByVal lpAsm + &H9 + 1     ' Call EbMode
        GetMem4 lpSetTimer - (lpAsm + &H23) - 5, ByVal lpAsm + &H23 + 1 ' Call SetTimer
        GetMem4 lpDefProc - (lpAsm + &H40) - 5, ByVal lpAsm + &H40 + 1  ' call DefSubclassProc
        GetMem4 lpFlag, ByVal lpAsm + &H2                               ' Cmp [flag], 0
        GetMem4 lpFlag, ByVal lpAsm + &H2C                              ' Inc [flag]
        
        lpAsm = lpAsm + &H48
        
    End If
    
    Dim lpMeth      As Long
    Dim vTable      As Long
    
    GetMem4 ByVal ObjPtr(Me), vTable
    GetMem4 ByVal vTable + WNDPROCINDEX * 4 + &H1C, lpMeth
    GetMem4 ObjPtr(Me), ByVal lpAsm + &H10                             ' Push Me
    GetMem4 lpMeth - (lpAsm + &H14) - 5, ByVal lpAsm + &H14 + 1        ' Call WndProc
        
    If isFirst Then
        
        lpDefProc = GetProcAddress(hInstUser32, "DefWindowProcW")
        
        cls.hInstance = App.hInstance
        cls.lpfnwndproc = lpDefProc
        cls.lpszClassName = StrPtr(Mp3Class)
        cls.cbSize = Len(cls)
        cls.cbWndExtra2 = 8
        
        If RegisterClassEx(cls) = 0 Then
  
            HeapDestroy hHeap
            Exit Sub

        End If
        
        hwnd = CreateWindowEx(0, StrPtr(Mp3Class), 0, 0, 0, 0, 0, 0, HWND_MESSAGE, 0, App.hInstance, ByVal 0&)
        If hwnd = 0 Then Exit Sub
        
        SaveWindowAndHeap hwnd, hHeap
        
        If inIDE Then Call SetWindowLong(hwnd, 0, lpFlag)
    
    End If
        
    If SetWindowSubclass(hwnd, lpWndProc, ObjPtr(Me), 0) = 0 Then Exit Sub
    
    SetWindowLong hwnd, GWL_USERDATA, GetWindowLong(hwnd, GWL_USERDATA) + 1
    
    init = True
    
End Sub

Private Sub Class_Terminate()
    Dim refCt   As Long
    
    If Not init Then Exit Sub
    
    refCt = GetWindowLong(hwnd, GWL_USERDATA)
    
    If refCt = 0 Then
    
        DestroyWindow hwnd
        HeapDestroy hHeap
        UnregisterClass StrPtr(Mp3Class), App.hInstance
        SaveWindowAndHeap 0, 0
        
    Else
        
        RemoveWindowSubclass hwnd, lpWndProc, ObjPtr(Me)
        SetWindowLong hwnd, GWL_USERDATA, refCt - 1
        HeapFree hHeap, HEAP_NO_SERIALIZE, ByVal lpWndProc
        
    End If
    
End Sub

Private Function MakeTrue(refBool As Boolean) As Boolean
    MakeTrue = True
    refBool = True
End Function



