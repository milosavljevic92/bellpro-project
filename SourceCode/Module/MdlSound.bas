Attribute VB_Name = "MdlSound"
Option Explicit
Dim player  As clsTrickMP3Player
Public Sub PlayMp3Sound()
If Not DebugMode = True Then On Error Resume Next
    If getDemoStatus = True Then Exit Sub
    If GetProfile("config", "037", 0, getConfigPath) = "0" Then Exit Sub
    If Not player.IsPlaying = True Then player.Play True
End Sub
Public Sub StopMp3Sound()
    If Not DebugMode = True Then On Error Resume Next
    If getDemoStatus = True Then Exit Sub
    If GetProfile("config", "037", 0, getConfigPath) = "0" Then Exit Sub
    player.StopPlaying
End Sub
Public Sub initPlayer()
On Error Resume Next
    If getDemoStatus = True Then Exit Sub
    Set player = New clsTrickMP3Player
    Dim dat() As Byte
    dat = LoadResData(101, "CUSTOM")
    player.Initialize VarPtr(dat(0)), UBound(dat) + 1, True
    Exit Sub
End Sub

