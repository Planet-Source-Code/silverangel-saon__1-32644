Attribute VB_Name = "Module1"
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Sub mmOpen(ByVal sTheFile As String)

    Dim nReturn As Long
    
    Dim sType As String
    
    If sAlias <> "" Then
        mmClose
    End If
    
    Select Case UCase$(Right$(sTheFile, 3))
        Case "WAV"
            sType = "Waveaudio"
        Case "AVI"
            sType = "AviVideo"
        Case "MID"
            sType = "Sequencer"
        Case "MP3"
            sType = "MPegVideo"
        Case Else
            Exit Sub
    End Select
    Randomize
    sAlias = Right$(sTheFile, 3) & Minute(Now) & Second(Now) & Int(1000 * Rnd + 1)
    
    If InStr(sTheFile, " ") Then sTheFile = Chr(34) & sTheFile & Chr(34)
    nReturn = mciSendString("Open " & sTheFile & " ALIAS " & sAlias & " TYPE " & sType & " wait", "", 0, 0)
    
End Sub

Public Sub mmClose()

    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Close " & sAlias, "", 0, 0)
    sAlias = ""
    sFilename = ""
    
End Sub

Public Sub mmPlay()

    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    
    If bWait Then
        nReturn = mciSendString("Play " & sAlias & " wait", "", 0, 0)
    Else
        nReturn = mciSendString("Play " & sAlias, "", 0, 0)
    End If
    
End Sub

Public Sub mmStop()

    Dim nReturn As Long
    
    If sAlias = "" Then Exit Sub
    
    nReturn = mciSendString("Stop " & sAlias, "", 0, 0)
    
End Sub

