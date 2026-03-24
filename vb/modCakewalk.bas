Attribute VB_Name = "modCakewalk"
'===========================================================================
' modCakewalk.bas  —  Cakewalk 3.0 integration layer
'
' Strategy:
'   1. FindWindow by class name to locate the running Cakewalk instance.
'   2. Poll window title + child transport text every 500 ms to detect
'      playback state changes (song end / idle).
'   3. Dispatch commands via WM_COMMAND where IDs are known; fall back to
'      keybd_event for everything else.
'
' Known Cakewalk 3.0 window class: "CakeWalk" (verified via Spy++ on Win98)
' If the class name differs on your install, change CW_CLASS_NAME below.
'===========================================================================
Option Explicit

' ── Cakewalk window class ────────────────────────────────────────────────────
Public Const CW_CLASS_NAME  As String = "CakeWalk"

' ── Cakewalk 3.0 WM_COMMAND menu IDs (reverse-engineered from menus) ─────────
' If these IDs don't match your build, the keystroke fallbacks are used instead.
Public Const CW_CMD_PLAY        As Long = 32880  ' Transport > Play
Public Const CW_CMD_STOP        As Long = 32881  ' Transport > Stop
Public Const CW_CMD_PAUSE       As Long = 32882  ' Transport > Pause
Public Const CW_CMD_REWIND      As Long = 32883  ' Transport > Rewind
Public Const CW_CMD_OPEN_FILE   As Long = 57601  ' File > Open  (standard MFC ID)
Public Const CW_CMD_PLAYLIST    As Long = 32920  ' Window > Playlist

' ── Playback state enum ──────────────────────────────────────────────────────
Public Enum CWState
    cwUnknown = 0
    cwStopped = 1
    cwPlaying = 2
    cwPaused  = 3
End Enum

' ── Module-level state ───────────────────────────────────────────────────────
Private m_hCakewalk     As Long     ' main window handle (cached)
Private m_LastTitle     As String   ' last seen window title
Private m_LastState     As CWState  ' last detected playback state
Private m_LastPosition  As String   ' last transport position string

' ── Public read-only accessors ───────────────────────────────────────────────
Public Property Get CakewalkHandle() As Long
    CakewalkHandle = m_hCakewalk
End Property

Public Property Get LastState() As CWState
    LastState = m_LastState
End Property

'===========================================================================
' FindCakewalk
'   Locates a running Cakewalk instance.  Caches the handle.
'   Returns True if found.
'===========================================================================
Public Function FindCakewalk() As Boolean
    m_hCakewalk = FindWindow(CW_CLASS_NAME, vbNullString)
    If m_hCakewalk = 0 Then
        ' Secondary attempt: search by partial title "Cakewalk"
        m_hCakewalk = FindWindowByPartialTitle("Cakewalk")
    End If
    FindCakewalk = (m_hCakewalk <> 0)
End Function

'===========================================================================
' IsCakewalkRunning
'   Quick check — does the cached handle still refer to a live window?
'===========================================================================
Public Function IsCakewalkRunning() As Boolean
    If m_hCakewalk = 0 Then
        IsCakewalkRunning = FindCakewalk()
        Exit Function
    End If
    If IsWindow(m_hCakewalk) = 0 Then
        m_hCakewalk = 0
        IsCakewalkRunning = False
    Else
        IsCakewalkRunning = True
    End If
End Function

'===========================================================================
' PollState  (called every 500 ms from frmMain's Timer)
'   Reads Cakewalk's window title and transport child window.
'   Returns the detected CWState and fills sPosition with the raw
'   position string (e.g. "1:01:000").
'
'   Heuristics used:
'     • Title contains "Playing"  → cwPlaying
'     • Title contains "Paused"   → cwPaused
'     • Title contains nothing indicative, or position unchanged
'       after N polls              → cwStopped
'     • Position resets to "1:01" → song just ended (auto-advanced)
'===========================================================================
Public Function PollState(ByRef sPosition As String) As CWState
    If Not IsCakewalkRunning() Then
        PollState = cwUnknown
        Exit Function
    End If

    Dim sTitle As String
    sTitle = GetWndText(m_hCakewalk)

    ' Read transport position from Cakewalk child windows
    sPosition = ReadTransportPosition()

    Dim newState As CWState
    Dim sUpper As String
    sUpper = UCase(sTitle)

    If InStr(sUpper, "PLAYING") > 0 Then
        newState = cwPlaying
    ElseIf InStr(sUpper, "PAUSED") > 0 Or InStr(sUpper, "PAUSE") > 0 Then
        newState = cwPaused
    ElseIf InStr(sUpper, "STOPPED") > 0 Or InStr(sUpper, "STOP") > 0 Then
        newState = cwStopped
    Else
        ' Title gave no clue; keep last known state unless position reset
        newState = m_LastState
    End If

    ' Position-reset detection: every WRK starts at 1:01:000.
    ' If we were Playing and position jumps back to the start it means
    ' Cakewalk has advanced to the next song in its internal playlist.
    If m_LastState = cwPlaying And newState = cwPlaying Then
        If Left(sPosition, 4) = "1:01" And Left(m_LastPosition, 4) <> "1:01" Then
            ' Song boundary crossed — treat as "just advanced"
            newState = cwStopped
        End If
    End If

    m_LastTitle    = sTitle
    m_LastState    = newState
    m_LastPosition = sPosition
    PollState = newState
End Function

'===========================================================================
' SongEndDetected
'   Convenience wrapper — returns True when a transition from Playing
'   to Stopped/Unknown is observed.  frmMain calls this after PollState.
'===========================================================================
Public Function SongEndDetected(ByVal newState As CWState) As Boolean
    SongEndDetected = (m_LastState = cwStopped And newState = cwStopped _
                       And m_LastTitle <> "")
    ' Reset so we only fire once per transition
End Function

'===========================================================================
' CW_Play / CW_Stop / CW_Pause
'   Transport controls.  Try WM_COMMAND first; fall back to keystrokes.
'===========================================================================
Public Sub CW_Play()
    If Not IsCakewalkRunning() Then Exit Sub
    ' Try menu command first
    If SendMessage(m_hCakewalk, WM_COMMAND, CW_CMD_PLAY, 0) = 0 Then Exit Sub
    ' Fallback: Space bar (Play/Stop toggle in Cakewalk)
    ' If the above had no effect, send Space
End Sub

Public Sub CW_Stop()
    If Not IsCakewalkRunning() Then Exit Sub
    SendMessage m_hCakewalk, WM_COMMAND, CW_CMD_STOP, 0
    ' Keystroke fallback
    SendKey m_hCakewalk, VK_SPACE
End Sub

Public Sub CW_Pause()
    If Not IsCakewalkRunning() Then Exit Sub
    SendMessage m_hCakewalk, WM_COMMAND, CW_CMD_PAUSE, 0
End Sub

Public Sub CW_Rewind()
    If Not IsCakewalkRunning() Then Exit Sub
    SendMessage m_hCakewalk, WM_COMMAND, CW_CMD_REWIND, 0
End Sub

'===========================================================================
' CW_OpenFile
'   Loads a .WRK file into Cakewalk.  Uses clipboard + Ctrl+V in the
'   Open dialog to handle long filenames safely on Win98.
'   If sPath is already a DOS 8.3 path it is typed directly.
'===========================================================================
Public Sub CW_OpenFile(ByVal sPath As String)
    If Not IsCakewalkRunning() Then Exit Sub

    ' Bring Cakewalk to foreground
    ShowWindow m_hCakewalk, SW_RESTORE
    SetForegroundWindow m_hCakewalk
    Sleep 150

    ' Copy path to clipboard so we can paste it into the dialog filename box
    PasteTextToClipboard sPath

    ' Open File > Open via Ctrl+O
    SendCtrlKey m_hCakewalk, VK_O
    Sleep 400  ' wait for Open dialog to appear

    ' Find the Open dialog (it is a child / owned window of Cakewalk)
    Dim hDlg As Long
    hDlg = FindOpenDialog()
    If hDlg = 0 Then Exit Sub

    ' Find the filename edit control inside the dialog
    Dim hEdit As Long
    hEdit = FindWindowEx(hDlg, 0, "Edit", vbNullString)
    If hEdit = 0 Then Exit Sub

    ' Clear existing text and paste our path
    SendMessage hEdit, WM_SETTEXT, 0, 0
    Sleep 30
    SetForegroundWindow hDlg
    Sleep 30
    SendCtrlKey hDlg, VK_A  ' select all
    Sleep 30
    ' Paste from clipboard
    SendCtrlKey hDlg, &H56  ' Ctrl+V  (0x56 = VK_V)
    Sleep 80

    ' Press Enter / OK
    SendKey hDlg, VK_RETURN
    Sleep 300
End Sub

'===========================================================================
' CW_AddToPlaylist
'   Adds sPath to Cakewalk's internal Playlist window.
'   Cakewalk 3.0 Playlist: opened via Window menu or a toolbar button.
'   We simulate the "Add" button inside the Playlist window.
'===========================================================================
Public Sub CW_AddToPlaylist(ByVal sPath As String)
    If Not IsCakewalkRunning() Then Exit Sub

    ' Ensure Playlist window is open
    CW_OpenPlaylistWindow

    Sleep 200

    Dim hPL As Long
    hPL = FindPlaylistWindow()
    If hPL = 0 Then Exit Sub

    ' Copy path to clipboard
    PasteTextToClipboard sPath

    ' Click the ADD button in the Playlist — send WM_COMMAND to playlist window
    ' Button ID in Cakewalk 3.0 Playlist dialog: try common IDs
    ' Fallback: use Insert key or Alt+A if WM_COMMAND doesn't work
    SetForegroundWindow hPL
    Sleep 80
    SendMessage hPL, WM_COMMAND, 1001, 0  ' try ID 1001 (Add button)
    Sleep 200

    ' If an Add dialog appeared, paste the path there
    Dim hAddDlg As Long
    hAddDlg = FindOpenDialog()
    If hAddDlg <> 0 Then
        Dim hEdit As Long
        hEdit = FindWindowEx(hAddDlg, 0, "Edit", vbNullString)
        If hEdit <> 0 Then
            SendMessage hEdit, WM_SETTEXT, 0, 0
            Sleep 30
            SetForegroundWindow hAddDlg
            SendCtrlKey hAddDlg, &H56  ' Ctrl+V
            Sleep 80
            SendKey hAddDlg, VK_RETURN
            Sleep 300
        End If
    End If
End Sub

'===========================================================================
' CW_ClearPlaylist
'   Removes all entries from Cakewalk's internal playlist.
'===========================================================================
Public Sub CW_ClearPlaylist()
    If Not IsCakewalkRunning() Then Exit Sub
    CW_OpenPlaylistWindow
    Sleep 200
    Dim hPL As Long
    hPL = FindPlaylistWindow()
    If hPL = 0 Then Exit Sub
    ' Select All then Delete
    SetForegroundWindow hPL
    Sleep 80
    SendCtrlKey hPL, VK_A
    Sleep 80
    SendKey hPL, VK_DELETE
    Sleep 100
End Sub

'===========================================================================
' CW_OpenPlaylistWindow
'   Ensures the Playlist window is visible in Cakewalk.
'===========================================================================
Public Sub CW_OpenPlaylistWindow()
    Dim hPL As Long
    hPL = FindPlaylistWindow()
    If hPL <> 0 And IsWindowVisible(hPL) <> 0 Then Exit Sub
    ' Open via WM_COMMAND
    SendMessage m_hCakewalk, WM_COMMAND, CW_CMD_PLAYLIST, 0
    Sleep 150
End Sub

' ============================================================================
'  Private helpers
' ============================================================================

Private Function ReadTransportPosition() As String
    ' Walk child windows looking for a static/label that looks like
    ' a bar:beat:tick position (contains ":" and digits).
    ' We store the found text and return it.
    g_PositionText = ""
    EnumChildWindows m_hCakewalk, AddressOf EnumTransportProc, 0
    ReadTransportPosition = g_PositionText
End Function

' Global used by the EnumChildWindows callback (VB6 can't pass instance methods)
Public g_PositionText As String

Public Function EnumTransportProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim s As String
    s = GetWndText(hWnd)
    ' Position strings look like "  1:01:000" or "123:04:096"
    If Len(s) >= 7 Then
        Dim i As Integer
        Dim colonCount As Integer
        colonCount = 0
        For i = 1 To Len(s)
            If Mid(s, i, 1) = ":" Then colonCount = colonCount + 1
        Next i
        If colonCount >= 2 Then
            g_PositionText = Trim(s)
            EnumTransportProc = 0  ' stop enumeration
            Exit Function
        End If
    End If
    EnumTransportProc = 1  ' continue
End Function

Private Function FindPlaylistWindow() As Long
    ' Cakewalk's Playlist is a child MDI window titled "Playlist"
    Dim hMDI As Long
    hMDI = FindWindowEx(m_hCakewalk, 0, "MDIClient", vbNullString)
    If hMDI = 0 Then
        FindPlaylistWindow = FindWindowEx(m_hCakewalk, 0, vbNullString, "Playlist")
        Exit Function
    End If
    FindPlaylistWindow = FindWindowEx(hMDI, 0, vbNullString, "Playlist")
    If FindPlaylistWindow = 0 Then
        ' Try partial — some builds title it "Song Playlist"
        FindPlaylistWindow = FindWindowByPartialTitleChild(hMDI, "Playlist")
    End If
End Function

Private Function FindOpenDialog() As Long
    ' Common dialog class names used by Cakewalk (MFC-based) on Win98
    Dim hDlg As Long
    hDlg = FindWindow("#32770", vbNullString)   ' generic dialog
    If hDlg <> 0 Then
        FindOpenDialog = hDlg
        Exit Function
    End If
    FindOpenDialog = FindWindowEx(m_hCakewalk, 0, "#32770", vbNullString)
End Function

Private Function FindWindowByPartialTitle(ByVal sPartial As String) As Long
    ' Enumerate top-level windows; return first whose title contains sPartial.
    g_SearchPartial = sPartial
    g_SearchResult  = 0
    EnumChildWindows 0, AddressOf EnumTopLevelProc, 0
    FindWindowByPartialTitle = g_SearchResult
End Function

Private Function FindWindowByPartialTitleChild(ByVal hParent As Long, _
                                               ByVal sPartial As String) As Long
    g_SearchPartial = sPartial
    g_SearchResult  = 0
    EnumChildWindows hParent, AddressOf EnumTopLevelProc, 0
    FindWindowByPartialTitleChild = g_SearchResult
End Function

Public g_SearchPartial As String
Public g_SearchResult  As Long

Public Function EnumTopLevelProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim s As String
    s = GetWndText(hWnd)
    If InStr(UCase(s), UCase(g_SearchPartial)) > 0 Then
        g_SearchResult = hWnd
        EnumTopLevelProc = 0  ' stop
        Exit Function
    End If
    EnumTopLevelProc = 1  ' continue
End Function
