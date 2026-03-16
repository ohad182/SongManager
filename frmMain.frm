VERSION 5.00
Begin VB.Form frmMain
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SongManager v1.0  —  Cakewalk 3.0 Playlist Wrapper"
   ClientHeight    =   8400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7200
   Icon            =   "NONE"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSearch
      BackColor       =   &H00202020&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   ""
      Top             =   120
      Width           =   5640
   End
   Begin VB.Label lblSearchHint
      BackColor       =   &H00404040&
      Caption         =   "Search:"
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   5820
      Top             =   180
      Width           =   735
   End
   Begin VB.ListBox lstSongs
      BackColor       =   &H00202020&
      ForeColor       =   &H0000FF00&
      Height          =   5400
      IntegralHeight  =   0
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   6960
   End
   ' ── Status bar ──────────────────────────────────────────────────────────
   Begin VB.Label lblStatus
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cakewalk: NOT FOUND"
      ForeColor       =   &H00FF4040&
      Height          =   300
      Left            =   120
      Top             =   6030
      Width           =   6960
   End
   ' ── Transport controls ──────────────────────────────────────────────────
   Begin VB.CommandButton btnPlay
      BackColor       =   &H00006000&
      Caption         =   "PLAY"
      Height          =   480
      Left            =   120
      Style           =   1  'Graphical
      Top             =   6420
      Width           =   1050
   End
   Begin VB.CommandButton btnStop
      BackColor       =   &H00600000&
      Caption         =   "STOP"
      Height          =   480
      Left            =   1260
      Style           =   1  'Graphical
      Top             =   6420
      Width           =   1050
   End
   Begin VB.CommandButton btnPause
      BackColor       =   &H00606000&
      Caption         =   "PAUSE"
      Height          =   480
      Left            =   2400
      Style           =   1  'Graphical
      Top             =   6420
      Width           =   1050
   End
   Begin VB.CommandButton btnNext
      BackColor       =   &H00404040&
      Caption         =   "NEXT >"
      Height          =   480
      Left            =   3540
      Style           =   1  'Graphical
      Top             =   6420
      Width           =   1050
   End
   ' ── Transpose stubs ─────────────────────────────────────────────────────
   Begin VB.CommandButton btnTrMinus
      BackColor       =   &H00404040&
      Caption         =   "Tr -"
      Height          =   480
      Left            =   4680
      Style           =   1  'Graphical
      Top             =   6420
      Width           =   750
   End
   Begin VB.Label lblTranspose
      Alignment       =   2  'Center
      BackColor       =   &H00202020&
      BorderStyle     =   1
      Caption         =   " 0"
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Left            =   5490
      Top             =   6420
      Width           =   480
   End
   Begin VB.CommandButton btnTrPlus
      BackColor       =   &H00404040&
      Caption         =   "Tr +"
      Height          =   480
      Left            =   6030
      Style           =   1  'Graphical
      Top             =   6420
      Width           =   750
   End
   ' ── Playlist management buttons ─────────────────────────────────────────
   Begin VB.CommandButton btnAdd
      BackColor       =   &H00404040&
      Caption         =   "Add"
      Height          =   390
      Left            =   120
      Style           =   1  'Graphical
      Top             =   7020
      Width           =   870
   End
   Begin VB.CommandButton btnDelete
      BackColor       =   &H00404040&
      Caption         =   "Delete"
      Height          =   390
      Left            =   1080
      Style           =   1  'Graphical
      Top             =   7020
      Width           =   870
   End
   Begin VB.CommandButton btnDelay
      BackColor       =   &H00404040&
      Caption         =   "Delay"
      Height          =   390
      Left            =   2040
      Style           =   1  'Graphical
      Top             =   7020
      Width           =   870
   End
   Begin VB.CommandButton btnClear
      BackColor       =   &H00404040&
      Caption         =   "Clear List"
      Height          =   390
      Left            =   3000
      Style           =   1  'Graphical
      Top             =   7020
      Width           =   1050
   End
   Begin VB.CommandButton btnLoadDB
      BackColor       =   &H00404040&
      Caption         =   "Load DB"
      Height          =   390
      Left            =   4140
      Style           =   1  'Graphical
      Top             =   7020
      Width           =   870
   End
   Begin VB.CommandButton btnLoadSet
      BackColor       =   &H00404040&
      Caption         =   "Load Set"
      Height          =   390
      Left            =   5100
      Style           =   1  'Graphical
      Top             =   7020
      Width           =   870
   End
   Begin VB.CommandButton btnSaveSet
      BackColor       =   &H00404040&
      Caption         =   "Save Set"
      Height          =   390
      Left            =   6060
      Style           =   1  'Graphical
      Top             =   7020
      Width           =   870
   End
   ' ── Font size controls ──────────────────────────────────────────────────
   Begin VB.CommandButton btnFontPlus
      BackColor       =   &H00404040&
      Caption         =   "Font +"
      Height          =   390
      Left            =   120
      Style           =   1  'Graphical
      Top             =   7530
      Width           =   870
   End
   Begin VB.CommandButton btnFontMinus
      BackColor       =   &H00404040&
      Caption         =   "Font -"
      Height          =   390
      Left            =   1080
      Style           =   1  'Graphical
      Top             =   7530
      Width           =   870
   End
   ' ── Full path toggle ────────────────────────────────────────────────────
   Begin VB.CheckBox chkFullPath
      BackColor       =   &H00404040&
      Caption         =   "Show Full Path"
      ForeColor       =   &H00C0C0C0&
      Height          =   390
      Left            =   2160
      Top             =   7560
      Width           =   1800
   End
   ' ── Polling timer ───────────────────────────────────────────────────────
   Begin VB.Timer tmrPoll
      Interval        =   500
      Left            =   4200
      Top             =   7560
   End
   ' ── Connection-check timer (slower) ─────────────────────────────────────
   Begin VB.Timer tmrConnect
      Interval        =   2000
      Left            =   5400
      Top             =   7560
   End
   ' ── Common dialog for file open ─────────────────────────────────────────
   Begin MSComDlg.CommonDialog dlgOpen
      Left            =   6600
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
' frmMain.frm  —  Main form: UI + event wiring
'
' Responsibilities:
'   • Render the song queue list with PLAYING / READY / index status labels
'   • Relay button clicks to modPlaylist and modCakewalk
'   • Drive the 500 ms polling timer → song-end detection → rolling buffer
'   • Handle search-as-you-type filtering
'   • Save / Load .SET files
'===========================================================================
Option Explicit

' ── Form-level state ─────────────────────────────────────────────────────────
Private m_TransposeSemitones As Integer   ' Tr+/Tr- counter (Phase 2 stub)
Private m_bShowFullPath      As Boolean
Private m_bUpdatingList      As Boolean   ' guard against re-entrant list refresh
Private m_PrevState          As CWState   ' last polled Cakewalk state

' ──────────────────────────────────────────────────────────────────────────────
'  FORM INIT / TERMINATE
' ──────────────────────────────────────────────────────────────────────────────

Private Sub Form_Load()
    m_TransposeSemitones = 0
    m_bShowFullPath      = False
    m_bUpdatingList      = False
    m_PrevState          = cwUnknown

    ' Initialise the playlist engine
    InitPlaylist

    ' Style the list font (large, monospaced for live-stage readability)
    lstSongs.Font.Name = "Courier New"
    lstSongs.Font.Size = 12

    ' Try to connect immediately
    UpdateConnectionStatus

    ' Start timers
    tmrPoll.Enabled    = True
    tmrConnect.Enabled = True

    lblTranspose.Caption = " 0"
    lblStatus.Caption    = "Ready  —  Queue empty"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    tmrPoll.Enabled    = False
    tmrConnect.Enabled = False
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  POLLING TIMER  (500 ms)
'  Detects Cakewalk state changes and drives the rolling buffer.
' ──────────────────────────────────────────────────────────────────────────────

Private Sub tmrPoll_Timer()
    If Not IsCakewalkRunning() Then Exit Sub

    Dim sPos     As String
    Dim newState As CWState
    newState = PollState(sPos)

    ' Update status bar with live position
    If newState = cwPlaying Then
        lblStatus.Caption  = "PLAYING  [" & sPos & "]  —  " & NowPlayingLabel()
        lblStatus.ForeColor = &H0000FF00&  ' green
    ElseIf newState = cwPaused Then
        lblStatus.Caption  = "PAUSED   [" & sPos & "]"
        lblStatus.ForeColor = &H0000FFFF&  ' yellow-cyan
    ElseIf newState = cwStopped Then
        lblStatus.Caption  = "STOPPED"
        lblStatus.ForeColor = &H00C0C0C0&
    End If

    ' Detect transition: Playing → Stopped  means song ended
    If m_PrevState = cwPlaying And newState = cwStopped Then
        OnSongAdvanced         ' modPlaylist: increment QueueIdx, load next slot
        RefreshList            ' redraw with updated PLAYING/READY labels
    End If

    m_PrevState = newState
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  CONNECTION TIMER  (2 s)
'  Checks if Cakewalk is still running; updates status bar.
' ──────────────────────────────────────────────────────────────────────────────

Private Sub tmrConnect_Timer()
    UpdateConnectionStatus
End Sub

Private Sub UpdateConnectionStatus()
    If IsCakewalkRunning() Then
        If m_PrevState <> cwPlaying And m_PrevState <> cwPaused Then
            lblStatus.Caption  = "Cakewalk: CONNECTED  —  " & _
                                 IIf(g_QueueSize > 0, _
                                     CStr(g_QueueSize) & " song(s) in queue", _
                                     "Queue empty")
            lblStatus.ForeColor = &H00C0C0C0&
        End If
    Else
        lblStatus.Caption  = "Cakewalk: NOT FOUND  —  Please start Cakewalk 3.0"
        lblStatus.ForeColor = &H00FF4040&  ' red
    End If
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  TRANSPORT BUTTONS
' ──────────────────────────────────────────────────────────────────────────────

Private Sub btnPlay_Click()
    If Not IsCakewalkRunning() Then
        MsgBox "Cakewalk is not running.", vbExclamation, "SongManager"
        Exit Sub
    End If
    If g_QueueSize = 0 Then
        MsgBox "The queue is empty. Add songs first.", vbInformation, "SongManager"
        Exit Sub
    End If

    If g_QueueIdx < 0 Then
        ' Nothing has been started yet — start from selected item (or top)
        Dim startIdx As Long
        startIdx = SelectedQueueIdx()
        If startIdx < 0 Then startIdx = 0
        StartPlayback startIdx
    Else
        CW_Play
    End If
    RefreshList
End Sub

Private Sub btnStop_Click()
    CW_Stop
    RefreshList
End Sub

Private Sub btnPause_Click()
    CW_Pause
End Sub

Private Sub btnNext_Click()
    SkipToNext
    RefreshList
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  TRANSPOSE STUBS  (Phase 2)
' ──────────────────────────────────────────────────────────────────────────────

Private Sub btnTrPlus_Click()
    If m_TransposeSemitones < 12 Then
        m_TransposeSemitones = m_TransposeSemitones + 1
        lblTranspose.Caption = FormatTranspose(m_TransposeSemitones)
    End If
    ' TODO Phase 2: apply transpose to MIDI output
End Sub

Private Sub btnTrMinus_Click()
    If m_TransposeSemitones > -12 Then
        m_TransposeSemitones = m_TransposeSemitones - 1
        lblTranspose.Caption = FormatTranspose(m_TransposeSemitones)
    End If
    ' TODO Phase 2: apply transpose to MIDI output
End Sub

Private Function FormatTranspose(ByVal n As Integer) As String
    If n > 0 Then
        FormatTranspose = "+" & CStr(n)
    Else
        FormatTranspose = CStr(n)
    End If
End Function

' ──────────────────────────────────────────────────────────────────────────────
'  PLAYLIST MANAGEMENT BUTTONS
' ──────────────────────────────────────────────────────────────────────────────

Private Sub btnAdd_Click()
    ' Browse for one or more .WRK files
    On Error Resume Next
    dlgOpen.Filter      = "Cakewalk Files (*.wrk)|*.WRK|All Files (*.*)|*.*"
    dlgOpen.FilterIndex = 1
    dlgOpen.Flags       = &H4 Or &H200  ' OFN_NOCHANGEDIR | OFN_ALLOWMULTISELECT
    dlgOpen.ShowOpen
    If Err.Number <> 0 Or dlgOpen.FileName = "" Then Exit Sub
    On Error GoTo 0

    Dim sFile As String
    sFile = dlgOpen.FileName

    ' CommonDialog returns multiple files separated by Null in multiselect mode,
    ' but VB6's comdlg returns only the first easily; iterate via FileTitle too.
    AddSong sFile, ""
    RefreshList
End Sub

Private Sub btnDelete_Click()
    Dim qi As Long
    qi = SelectedQueueIdx()
    If qi < 0 Then Exit Sub
    If qi = g_QueueIdx Then
        If MsgBox("Delete the currently playing song?", vbYesNo Or vbQuestion, _
                  "SongManager") = vbNo Then Exit Sub
    End If
    RemoveSong qi
    RefreshList
End Sub

Private Sub btnDelay_Click()
    Dim qi As Long
    qi = SelectedQueueIdx()
    If qi < 0 Then Exit Sub

    Dim sInput As String
    sInput = InputBox("Enter pre-play delay in seconds for:" & vbCrLf & _
                      g_Queue(qi).DisplayName, "Set Delay", _
                      CStr(GetDelay(qi)))
    If sInput = "" Then Exit Sub
    If IsNumeric(sInput) Then
        SetDelay qi, CInt(sInput)
    End If
End Sub

Private Sub btnClear_Click()
    If g_QueueSize > 0 Then
        If MsgBox("Clear the entire queue?", vbYesNo Or vbQuestion, _
                  "SongManager") = vbNo Then Exit Sub
    End If
    ClearQueue
    RefreshList
End Sub

Private Sub btnLoadDB_Click()
    On Error Resume Next
    dlgOpen.Filter      = "Database Files (*.csv;*.xls)|*.CSV;*.XLS|CSV Files (*.csv)|*.CSV|Excel Files (*.xls)|*.XLS"
    dlgOpen.FilterIndex = 1
    dlgOpen.Flags       = &H4  ' OFN_NOCHANGEDIR
    dlgOpen.ShowOpen
    If Err.Number <> 0 Or dlgOpen.FileName = "" Then Exit Sub
    On Error GoTo 0

    Dim nAdded As Long
    Dim sError As String
    If DetectAndLoad(dlgOpen.FileName, nAdded, sError) Then
        MsgBox "Loaded " & nAdded & " song(s) from database.", vbInformation, "SongManager"
        RefreshList
    Else
        MsgBox "Import failed:" & vbCrLf & sError, vbCritical, "SongManager"
    End If
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  SAVE / LOAD SET
' ──────────────────────────────────────────────────────────────────────────────

Private Sub btnSaveSet_Click()
    On Error Resume Next
    dlgOpen.Filter      = "SongManager Set Files (*.set)|*.SET|All Files (*.*)|*.*"
    dlgOpen.FilterIndex = 1
    dlgOpen.Flags       = &H4  ' OFN_NOCHANGEDIR
    dlgOpen.ShowSave
    If Err.Number <> 0 Or dlgOpen.FileName = "" Then Exit Sub
    On Error GoTo 0

    Dim sPath As String
    sPath = dlgOpen.FileName
    If UCase(Right(sPath, 4)) <> ".SET" Then sPath = sPath & ".SET"

    Dim nErr As Long
    nErr = SaveSet(sPath)
    If nErr = 0 Then
        MsgBox "Set saved: " & sPath, vbInformation, "SongManager"
    Else
        MsgBox "Save failed.", vbCritical, "SongManager"
    End If
End Sub

Private Sub btnLoadSet_Click()
    On Error Resume Next
    dlgOpen.Filter      = "SongManager Set Files (*.set)|*.SET|All Files (*.*)|*.*"
    dlgOpen.FilterIndex = 1
    dlgOpen.Flags       = &H4
    dlgOpen.ShowOpen
    If Err.Number <> 0 Or dlgOpen.FileName = "" Then Exit Sub
    On Error GoTo 0

    Dim nLoaded As Long
    Dim sError  As String
    If LoadSet(dlgOpen.FileName, nLoaded, sError) Then
        MsgBox "Loaded " & nLoaded & " song(s) from set.", vbInformation, "SongManager"
        RefreshList
    Else
        MsgBox "Load failed:" & vbCrLf & sError, vbCritical, "SongManager"
    End If
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  FONT SIZE CONTROLS
' ──────────────────────────────────────────────────────────────────────────────

Private Sub btnFontPlus_Click()
    If lstSongs.Font.Size < 36 Then
        lstSongs.Font.Size = lstSongs.Font.Size + 2
    End If
End Sub

Private Sub btnFontMinus_Click()
    If lstSongs.Font.Size > 6 Then
        lstSongs.Font.Size = lstSongs.Font.Size - 2
    End If
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  FULL PATH TOGGLE
' ──────────────────────────────────────────────────────────────────────────────

Private Sub chkFullPath_Click()
    m_bShowFullPath = (chkFullPath.Value = 1)
    RefreshList
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  SEARCH BOX  —  search-as-you-type
' ──────────────────────────────────────────────────────────────────────────────

Private Sub txtSearch_Change()
    BuildFilter txtSearch.Text
    RefreshList
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            ' Play the currently selected (or first filtered) song
            Dim qi As Long
            qi = SelectedQueueIdx()
            If qi >= 0 Then
                If IsCakewalkRunning() Then
                    StartPlayback qi
                    RefreshList
                    txtSearch.Text = ""
                    BuildFilter ""
                End If
            End If
        Case vbKeyDown
            ' Move selection down in list
            If lstSongs.ListIndex < lstSongs.ListCount - 1 Then
                lstSongs.ListIndex = lstSongs.ListIndex + 1
            End If
        Case vbKeyUp
            ' Move selection up in list
            If lstSongs.ListIndex > 0 Then
                lstSongs.ListIndex = lstSongs.ListIndex - 1
            End If
        Case vbKeyEscape
            txtSearch.Text = ""
            BuildFilter ""
            RefreshList
    End Select
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  LIST BOX  —  double-click to start playback
' ──────────────────────────────────────────────────────────────────────────────

Private Sub lstSongs_DblClick()
    Dim qi As Long
    qi = SelectedQueueIdx()
    If qi < 0 Then Exit Sub
    If Not IsCakewalkRunning() Then
        MsgBox "Cakewalk is not running.", vbExclamation, "SongManager"
        Exit Sub
    End If
    StartPlayback qi
    RefreshList
End Sub

' ──────────────────────────────────────────────────────────────────────────────
'  HELPERS
' ──────────────────────────────────────────────────────────────────────────────

'  RefreshList
'   Rebuilds lstSongs from the current queue / filter state.
'   Preserves the selected item by queue index.
Public Sub RefreshList()
    If m_bUpdatingList Then Exit Sub
    m_bUpdatingList = True

    Dim prevQIdx As Long
    prevQIdx = SelectedQueueIdx()

    lstSongs.Clear

    Dim count As Long
    Dim i     As Long

    If g_FilterActive Then
        count = g_FilteredSize
    Else
        count = g_QueueSize
    End If

    For i = 0 To count - 1
        Dim qi As Long
        If g_FilterActive Then
            qi = g_FilteredIdx(i)
        Else
            qi = i
        End If
        lstSongs.AddItem GetDisplayLabel(qi, m_bShowFullPath)
    Next i

    ' Restore selection
    If prevQIdx >= 0 Then
        For i = 0 To lstSongs.ListCount - 1
            Dim mappedQ As Long
            mappedQ = FilteredToQueueIdx(i)
            If mappedQ = prevQIdx Then
                lstSongs.ListIndex = i
                Exit For
            End If
        Next i
    End If

    ' If nothing selected but queue has items, select PLAYING row
    If lstSongs.ListIndex < 0 And g_QueueIdx >= 0 Then
        For i = 0 To lstSongs.ListCount - 1
            If FilteredToQueueIdx(i) = g_QueueIdx Then
                lstSongs.ListIndex = i
                Exit For
            End If
        Next i
    End If

    m_bUpdatingList = False
End Sub

'  SelectedQueueIdx
'   Returns the master-queue index for the currently highlighted list row,
'   or -1 if nothing is selected.
Private Function SelectedQueueIdx() As Long
    If lstSongs.ListIndex < 0 Then
        SelectedQueueIdx = -1
        Exit Function
    End If
    SelectedQueueIdx = FilteredToQueueIdx(lstSongs.ListIndex)
End Function

'  NowPlayingLabel
'   Returns the display name of the currently playing song for the status bar.
Private Function NowPlayingLabel() As String
    If g_QueueIdx >= 0 And g_QueueIdx < g_QueueSize Then
        NowPlayingLabel = g_Queue(g_QueueIdx).DisplayName
        If g_Queue(g_QueueIdx).Artist <> "" Then
            NowPlayingLabel = NowPlayingLabel & "  (" & g_Queue(g_QueueIdx).Artist & ")"
        End If
    Else
        NowPlayingLabel = ""
    End If
End Function
