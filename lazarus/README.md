======================================================================
SongManager v1.0  —  Lazarus / Free Pascal Port
Target: Windows 10 / 11   Toolchain: Lazarus + Free Pascal (free)
======================================================================

TOOLCHAIN INSTALL (one-time)
-----------------------------
  1. Download Lazarus from https://lazarus-ide.org
     Pick: Windows 64-bit installer (includes Free Pascal compiler)
  2. Run the installer — default path C:\lazarus
  3. No registration or license key required.

SOURCE FILES
------------
  SongManager.lpi    Lazarus project file (open this in the IDE)
  SongManager.lpr    Program entry point
  uAPI.pas           Win32 API helpers (clipboard, keybd_event, etc.)
  uCakewalk.pas      Cakewalk window detection & command dispatch
  uPlaylist.pas      Unlimited queue + rolling 2-song buffer engine
  uDB.pas            CSV / Excel (.xls) database importer
  uSet.pas           Save / Load .SET playlist files
  uMain.pas          Main form — UI and all event handlers
  uMain.lfm          Form layout definition

BUILD — IDE
-----------
  1. Open SongManager.lpi in Lazarus
  2. Press F9  (Run)  or  Shift+F9  (Build only)
  3. Output: SongManager.exe in this folder

BUILD — COMMAND LINE
--------------------
  cd C:\Users\ohadc\SongManager\lazarus

  Debug build (default):
    C:\lazarus\lazbuild.exe SongManager.lpi

  Release build (optimised, no debug info):
    C:\lazarus\lazbuild.exe --build-mode=Release SongManager.lpi

  Output binary: SongManager.exe
  Intermediate files: lib\x86_64-win64\  (safe to delete)

RUNTIME REQUIREMENTS
--------------------
  • Windows 10 / 11  (Win32/64 native — no .NET required)
  • Cakewalk Professional 3.0 must be running before clicking Play
  • For Excel import: Microsoft Jet 4.0 OLEDB / MDAC 2.6+
    (present on most Windows installs; free download if missing)

DIFFERENCES FROM VB6 VERSION
------------------------------
  • No COMDLG32.OCX dependency — uses native Lazarus dialogs
  • Builds and runs on modern Windows without compatibility mode
  • Excel import still uses ADO / Jet 4.0 (same as VB6 version)
  • Transpose (Tr+/Tr-) remains a Phase 2 stub, same as VB6

======================================================================
