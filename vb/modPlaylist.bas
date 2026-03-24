Attribute VB_Name = "modPlaylist"
'===========================================================================
' modPlaylist.bas  —  Unlimited queue + rolling 2-song buffer logic
'
' Rolling Buffer concept:
'   The external Queue() holds every song the user has queued.
'   Only 2 songs are ever loaded into Cakewalk at one time:
'
'     Slot 1  [PLAYING]  Queue(g_QueueIdx)
'     Slot 2  [READY]    Queue(g_QueueIdx + 1)
'
'   When Cakewalk signals song-end / auto-advance:
'     1. g_QueueIdx is incremented (slot 2 becomes slot 1).
'     2. Queue(g_QueueIdx + 1) is loaded into Cakewalk slot 2.
'     3. UI is refreshed.
'===========================================================================
Option Explicit

' ── Song data record ─────────────────────────────────────────────────────────
Public Type SongEntry
    DisplayName As String   ' bare filename, no extension, no path
    FullPath    As String   ' full OS path or DOS-shortened path as stored in DB
    Artist      As String   ' from DB; may be empty
End Type

' ── The unlimited master queue ───────────────────────────────────────────────
Public  g_Queue()    As SongEntry   ' dynamic array, 0-based
Public  g_QueueSize  As Long        ' number of entries currently in Queue
Public  g_QueueIdx   As Long        ' index of the currently PLAYING song (-1 = none)

' ── Filtered-view support (search) ──────────────────────────────────────────
' When search is active, g_FilteredIdx() maps filtered positions → g_Queue indexes
Public  g_FilteredIdx()  As Long
Public  g_FilteredSize   As Long
Public  g_FilterActive   As Boolean

' ── Delay support (per-song pre-play delay in seconds) ───────────────────────
' Stored parallel to the queue; default 0
Private m_Delay()  As Integer

'===========================================================================
' InitPlaylist
'   Must be called once at startup (or after Clear).
'===========================================================================
Public Sub InitPlaylist()
    g_QueueSize   = 0
    g_QueueIdx    = -1
    g_FilterActive = False
    g_FilteredSize = 0
    ReDim g_Queue(0)
    ReDim m_Delay(0)
    ReDim g_FilteredIdx(0)
End Sub

'===========================================================================
' AddSong
'   Appends a SongEntry to the master queue.
'   Returns the new 0-based index of the added entry.
'===========================================================================
Public Function AddSong(ByVal sFullPath As String, _
                        ByVal sArtist   As String) As Long
    ' Expand arrays
    If g_QueueSize = 0 Then
        ReDim g_Queue(0)
        ReDim m_Delay(0)
    Else
        ReDim Preserve g_Queue(g_QueueSize)
        ReDim Preserve m_Delay(g_QueueSize)
    End If

    With g_Queue(g_QueueSize)
        .FullPath    = sFullPath
        .Artist      = sArtist
        .DisplayName = StripToDisplayName(sFullPath)
    End With
    m_Delay(g_QueueSize) = 0

    AddSong = g_QueueSize
    g_QueueSize = g_QueueSize + 1
End Function

'===========================================================================
' RemoveSong
'   Removes the entry at queueIndex from the master queue.
'   Adjusts g_QueueIdx if needed.
'===========================================================================
Public Sub RemoveSong(ByVal queueIndex As Long)
    If queueIndex < 0 Or queueIndex >= g_QueueSize Then Exit Sub

    Dim i As Long
    For i = queueIndex To g_QueueSize - 2
        g_Queue(i) = g_Queue(i + 1)
        m_Delay(i) = m_Delay(i + 1)
    Next i
    g_QueueSize = g_QueueSize - 1
    If g_QueueSize > 0 Then
        ReDim Preserve g_Queue(g_QueueSize - 1)
        ReDim Preserve m_Delay(g_QueueSize - 1)
    End If

    ' Adjust current play index
    If queueIndex < g_QueueIdx Then
        g_QueueIdx = g_QueueIdx - 1
    ElseIf queueIndex = g_QueueIdx Then
        ' Removed the playing song; clamp
        If g_QueueIdx >= g_QueueSize Then g_QueueIdx = g_QueueSize - 1
    End If
End Sub

'===========================================================================
' ClearQueue
'   Empties the entire queue and resets state.
'===========================================================================
Public Sub ClearQueue()
    InitPlaylist
End Sub

'===========================================================================
' SetDelay
'   Sets a pre-play delay (seconds) for the song at queueIndex.
'===========================================================================
Public Sub SetDelay(ByVal queueIndex As Long, ByVal nSeconds As Integer)
    If queueIndex >= 0 And queueIndex < g_QueueSize Then
        m_Delay(queueIndex) = nSeconds
    End If
End Sub

Public Function GetDelay(ByVal queueIndex As Long) As Integer
    If queueIndex >= 0 And queueIndex < g_QueueSize Then
        GetDelay = m_Delay(queueIndex)
    Else
        GetDelay = 0
    End If
End Function

'===========================================================================
' StartPlayback
'   Initialises the rolling buffer starting at queueIndex.
'   Loads song[idx] and song[idx+1] into Cakewalk.
'   Returns False if there's nothing to play.
'===========================================================================
Public Function StartPlayback(ByVal queueIndex As Long) As Boolean
    If g_QueueSize = 0 Then
        StartPlayback = False
        Exit Function
    End If
    If queueIndex < 0 Then queueIndex = 0
    If queueIndex >= g_QueueSize Then
        StartPlayback = False
        Exit Function
    End If

    g_QueueIdx = queueIndex

    ' Clear Cakewalk's playlist and load the first 2 slots
    CW_ClearPlaylist
    Sleep 200

    CW_AddToPlaylist g_Queue(g_QueueIdx).FullPath
    If g_QueueIdx + 1 < g_QueueSize Then
        CW_AddToPlaylist g_Queue(g_QueueIdx + 1).FullPath
    End If

    Sleep 150
    CW_Play

    StartPlayback = True
End Function

'===========================================================================
' OnSongAdvanced
'   Called by frmMain's polling timer when a song-end transition is detected.
'   Advances the rolling buffer by one position:
'     • old slot 2 becomes slot 1 (already in Cakewalk)
'     • load new slot 2 from queue
'===========================================================================
Public Sub OnSongAdvanced()
    If g_QueueIdx < 0 Then Exit Sub

    g_QueueIdx = g_QueueIdx + 1

    ' If still within bounds, push the next song into Cakewalk slot 2
    Dim nextIdx As Long
    nextIdx = g_QueueIdx + 1
    If nextIdx < g_QueueSize Then
        CW_AddToPlaylist g_Queue(nextIdx).FullPath
    End If
    ' If g_QueueIdx has gone past the end, playback is naturally finished
End Sub

'===========================================================================
' SkipToNext
'   User-initiated skip.  Stops current, advances, re-arms buffer.
'===========================================================================
Public Sub SkipToNext()
    If g_QueueIdx + 1 >= g_QueueSize Then Exit Sub
    CW_Stop
    Sleep 150
    StartPlayback g_QueueIdx + 1
End Sub

'===========================================================================
' IsPlaying / IsReady helpers
'===========================================================================
Public Function IsPlayingIndex(ByVal queueIndex As Long) As Boolean
    IsPlayingIndex = (queueIndex = g_QueueIdx)
End Function

Public Function IsReadyIndex(ByVal queueIndex As Long) As Boolean
    IsReadyIndex = (g_QueueIdx >= 0 And queueIndex = g_QueueIdx + 1)
End Function

'===========================================================================
' GetDisplayLabel
'   Returns the status prefix + display name for the list box row at
'   queueIndex, e.g.  "► PLAYING  Simi Ner"  or  "  2  Shir Ahava"
'===========================================================================
Public Function GetDisplayLabel(ByVal queueIndex As Long, _
                                ByVal bShowFullPath As Boolean) As String
    If queueIndex < 0 Or queueIndex >= g_QueueSize Then
        GetDisplayLabel = ""
        Exit Function
    End If

    Dim sName As String
    If bShowFullPath Then
        sName = g_Queue(queueIndex).FullPath
    Else
        sName = g_Queue(queueIndex).DisplayName
    End If

    If IsPlayingIndex(queueIndex) Then
        GetDisplayLabel = Chr(16) & " PLAYING  " & sName  ' Chr(16) = ►
    ElseIf IsReadyIndex(queueIndex) Then
        GetDisplayLabel = "  READY   " & sName
    Else
        Dim nPos As Long
        nPos = queueIndex - g_QueueIdx
        If nPos < 0 Then nPos = queueIndex + 1
        GetDisplayLabel = "  " & Format(nPos, "###") & "       " & sName
    End If
End Function

'===========================================================================
' BuildFilter
'   Populates g_FilteredIdx with the subset of g_Queue entries whose
'   DisplayName or Artist contains sSearch (case-insensitive).
'   Pass "" to clear the filter.
'===========================================================================
Public Sub BuildFilter(ByVal sSearch As String)
    If Trim(sSearch) = "" Then
        g_FilterActive = False
        g_FilteredSize = 0
        Exit Sub
    End If

    g_FilterActive = True
    Dim sUp As String
    sUp = UCase(sSearch)

    ReDim g_FilteredIdx(g_QueueSize)   ' over-allocate
    g_FilteredSize = 0

    Dim i As Long
    For i = 0 To g_QueueSize - 1
        If InStr(UCase(g_Queue(i).DisplayName), sUp) > 0 Or _
           InStr(UCase(g_Queue(i).Artist), sUp) > 0 Or _
           InStr(UCase(g_Queue(i).FullPath), sUp) > 0 Then
            g_FilteredIdx(g_FilteredSize) = i
            g_FilteredSize = g_FilteredSize + 1
        End If
    Next i
End Sub

'===========================================================================
' FilteredToQueueIdx
'   Translates a filtered-list position to the master queue index.
'===========================================================================
Public Function FilteredToQueueIdx(ByVal filteredPos As Long) As Long
    If Not g_FilterActive Then
        FilteredToQueueIdx = filteredPos
    ElseIf filteredPos >= 0 And filteredPos < g_FilteredSize Then
        FilteredToQueueIdx = g_FilteredIdx(filteredPos)
    Else
        FilteredToQueueIdx = -1
    End If
End Function

' ============================================================================
'  Private helpers
' ============================================================================

'===========================================================================
' StripToDisplayName
'   "C:\ISRAELI2\ROCK\SIMI_NER.WRK" → "SIMI_NER"
'   Handles both backslash paths and DOS 8.3 names.
'===========================================================================
Public Function StripToDisplayName(ByVal sPath As String) As String
    ' Extract filename part
    Dim parts() As String
    parts = Split(sPath, "\")
    Dim sFile As String
    sFile = parts(UBound(parts))

    ' Strip extension
    Dim dotPos As Long
    dotPos = InStrRev(sFile, ".")
    If dotPos > 1 Then
        sFile = Left(sFile, dotPos - 1)
    End If

    ' Replace underscores with spaces for readability
    StripToDisplayName = Replace(sFile, "_", " ")
End Function

' InStrRev polyfill for VB6 (already built-in, but kept for clarity)
' VB6 has InStrRev natively from VB6 SP3+ — no polyfill needed.
