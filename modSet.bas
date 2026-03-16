Attribute VB_Name = "modSet"
'===========================================================================
' modSet.bas  —  Save / Load playlist sets
'
' File format  (.SET)  — plain text, one record per line:
'
'   # SongManager Set File
'   # Saved: 2026-03-16
'   C:\ISRAELI2\ROCK\SIMI_NER.WRK|Simi Ner|0
'   C:\HAVAAL~1\POP\SHIR.WRK||5
'
' Fields:  FullPath | Artist | DelaySeconds
' Lines starting with # are comments and are ignored on load.
'===========================================================================
Option Explicit

Private Const SET_SIGNATURE As String = "# SongManager Set File"
Private Const FIELD_SEP     As String = "|"

'===========================================================================
' SaveSet
'   Writes the current g_Queue to sFilePath.
'   Returns 0 on success, nonzero on error (Err.Number).
'===========================================================================
Public Function SaveSet(ByVal sFilePath As String) As Long
    SaveSet = 0
    On Error GoTo ErrSave

    Dim iFile As Integer
    iFile = FreeFile
    Open sFilePath For Output As #iFile

    Print #iFile, SET_SIGNATURE
    Print #iFile, "# Saved: " & Format(Now, "YYYY-MM-DD hh:mm")
    Print #iFile, "# Songs: " & CStr(g_QueueSize)
    Print #iFile, ""

    Dim i As Long
    For i = 0 To g_QueueSize - 1
        Print #iFile, g_Queue(i).FullPath & FIELD_SEP & _
                      g_Queue(i).Artist   & FIELD_SEP & _
                      CStr(GetDelay(i))
    Next i

    Close #iFile
    Exit Function

ErrSave:
    SaveSet = Err.Number
    On Error GoTo 0
    Close #iFile
End Function

'===========================================================================
' LoadSet
'   Reads a .SET file and appends its entries to the current queue.
'   Does NOT clear the queue first — caller may call ClearQueue() beforehand.
'   Returns True on success; fills sError on failure.
'===========================================================================
Public Function LoadSet(ByVal sFilePath As String, _
                         ByRef nLoaded   As Long, _
                         ByRef sError    As String) As Boolean
    nLoaded = 0
    sError  = ""

    If Len(Dir(sFilePath)) = 0 Then
        sError = "File not found: " & sFilePath
        LoadSet = False
        Exit Function
    End If

    Dim iFile  As Integer
    Dim sLine  As String
    Dim nLine  As Long

    On Error GoTo ErrLoad
    iFile = FreeFile
    Open sFilePath For Input As #iFile

    Do While Not EOF(iFile)
        Line Input #iFile, sLine
        nLine = nLine + 1
        sLine = Trim(sLine)

        ' Skip blanks and comments
        If Len(sLine) = 0 Or Left(sLine, 1) = "#" Then GoTo NextSetLine

        Dim parts() As String
        parts = Split(sLine, FIELD_SEP)

        Dim sPath   As String
        Dim sArtist As String
        Dim nDelay  As Integer

        sPath = Trim(parts(0))
        If UBound(parts) >= 1 Then sArtist = Trim(parts(1)) Else sArtist = ""
        If UBound(parts) >= 2 Then
            If IsNumeric(parts(2)) Then nDelay = CInt(parts(2))
        End If

        ' Validate: must be a .WRK path
        If UCase(Right(sPath, 4)) <> ".WRK" Then GoTo NextSetLine
        If InStr(sPath, "\") = 0 Then GoTo NextSetLine

        Dim qi As Long
        qi = AddSong(sPath, sArtist)
        SetDelay qi, nDelay
        nLoaded = nLoaded + 1

NextSetLine:
    Loop

    Close #iFile
    LoadSet = True
    Exit Function

ErrLoad:
    sError = "Set load error on line " & nLine & ": " & Err.Description
    On Error GoTo 0
    Close #iFile
    LoadSet = False
End Function
