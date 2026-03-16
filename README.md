======================================================================
SongManager v1.0  —  Cakewalk 3.0 Playlist Wrapper
Target: Windows 98 / ME   Language: VB6 SP6
======================================================================

OVERVIEW
--------
SongManager is a standalone Win32 utility that bypasses Cakewalk 3.0's
128-song limit and 8-character filename restriction.  It manages an
unlimited external queue and maintains a rolling 2-song buffer inside
Cakewalk's native playlist so transitions are instantaneous.


PROJECT FILES
-------------
  SongManager.vbp    VB6 project file (open this in VB6 IDE)
  frmMain.frm        Main form (UI + all event handlers)
  modAPI.bas         Win32 API declarations
  modCakewalk.bas    Cakewalk window detection & command dispatch
  modPlaylist.bas    Queue engine + rolling buffer state machine
  modDB.bas          CSV / Excel database importer
  modSet.bas         Save / Load .SET playlist files


BUILD REQUIREMENTS
------------------
  • Visual Basic 6 SP6 (or Borland/etc. VB6-compatible environment)
  • COMDLG32.OCX  must be registered (included with Win98 / VB6 runtime)
  • For Excel import: MDAC 2.6+ and Microsoft Jet 4.0 OLEDB provider
    (free download from Microsoft; already present on Win98 SE+)


RUNTIME REQUIREMENTS
--------------------
  • Windows 98 / ME  (Win32 only — no .NET required)
  • Cakewalk Professional 3.0 must already be running before clicking Play
  • VB6 runtime (MSVBVM60.DLL — distributed with VB6 apps)


HOW IT WORKS — ROLLING BUFFER
------------------------------
  1. Start Cakewalk 3.0 first; leave it running.
  2. Open SongManager.  The status bar will show "Cakewalk: CONNECTED".
  3. Load songs via:
       • [Load DB]   — import a CSV or Excel song database
       • [Add]       — browse for individual .WRK files
       • [Load Set]  — reload a previously saved .SET file
  4. Double-click a song (or select + press Enter) to start playback.
  5. SongManager loads 2 songs into Cakewalk's playlist:
         Slot 1  [PLAYING]  current song
         Slot 2  [READY]    next song
  6. When the current song ends (detected via window title polling),
     SongManager automatically loads the next queued song into slot 2.


CONTROLS
--------
  PLAY / STOP / PAUSE   Transport controls sent to Cakewalk
  NEXT >                Skip current song; reload buffer
  Tr- / Tr+             Transpose semitone counter (Phase 2 stub — no
                        audio effect yet; wired in a future release)
  Add                   Browse and add a .WRK file to the queue
  Delete                Remove the selected song from the queue
  Delay                 Set a pre-play pause (seconds) before a song
  Clear List            Empty the entire queue
  Load DB               Import song database (CSV or .XLS)
  Load Set              Load a .SET file into the queue
  Save Set              Save current queue to a .SET file
  Font + / Font -       Adjust list font size (for stage visibility)
  Show Full Path        Toggle display of full file paths in the list
  Search box            Type to instantly filter — press Enter to play
                        the highlighted result; Esc to clear


SEARCH
------
  Just start typing in the search box at the top.  The list filters
  instantly across SongName, Artist, and FilePath.
  • Enter      — play the selected (first) result immediately
  • Esc        — clear search and show full list
  • Up/Down    — move selection within filtered results


DATABASE FORMAT (CSV)
---------------------
  Plain text, UTF-8 or ANSI, comma-separated:
  SongName , Artist , FilePath

  Example:
    Simi Ner , Various , C:\ISRAELI2\ROCK\SIMI_NER.WRK
    Shir Ahava , , C:\HAVAAL~1\POP\SHIR.WRK
    "Cohen, Leonard" , Cohen , C:\ROCK\SUZANNE.WRK

  • First line may be a header row (auto-detected if FilePath column
    contains no backslash).
  • Lines starting with # are ignored.
  • Only .WRK files are imported.


DATABASE FORMAT (Excel .XLS)
-----------------------------
  Column headers (case-insensitive):
    SongName  |  Artist  |  FilePath
  Data starts on row 2 (header row required).
  Requires MDAC 2.6+ and Jet 4.0 OLEDB provider.


SET FILE FORMAT (.SET)
----------------------
  Plain text, one song per line:
    FullPath | Artist | DelaySeconds

  Example:
    C:\ROCK\SONG1.WRK|Led Zeppelin|0
    C:\POP\SONG2.WRK||3


CAKEWALK INTEGRATION NOTES
---------------------------
  SongManager finds Cakewalk by window class name "CakeWalk".
  If your Cakewalk install uses a different class (check with Spy++),
  edit the CW_CLASS_NAME constant in modCakewalk.bas.

  WM_COMMAND IDs used (CW 3.0):
    CW_CMD_PLAY     = 32880
    CW_CMD_STOP     = 32881
    CW_CMD_PAUSE    = 32882
    CW_CMD_PLAYLIST = 32920
  If these don't match your build, SongManager falls back to
  keyboard injection (keybd_event) automatically.

  Long filenames are passed to Cakewalk via the Windows clipboard
  (Ctrl+V into the Open dialog) to avoid 8.3 path limitations.


KNOWN LIMITATIONS / PHASE 2 TODOs
-----------------------------------
  • MIDI Transpose (Tr+/Tr-) display works but audio effect is a stub.
    Phase 2 will add a virtual MIDI loopback driver to re-pitch output.
  • Multi-select in the Add dialog returns only the first file; add
    songs one at a time or use Load DB for bulk import.
  • Cakewalk WM_COMMAND IDs may need adjustment for different CW 3.0
    builds; use a resource editor to verify.

======================================================================
