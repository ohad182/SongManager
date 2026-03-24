======================================================================
SongManager v1.0  —  Visual Basic 6 (Original)
Target: Windows 98 / ME   Language: VB6 SP6
======================================================================

TOOLCHAIN REQUIREMENT
---------------------
  Visual Basic 6 SP6  (commercial product, no longer sold by Microsoft)

  Obtaining VB6:
    • MSDN / Visual Studio subscription archive at
      https://my.visualstudio.com/Downloads  (search "Visual Studio 6")
    • Used physical CD (Visual Studio 6.0) available on eBay ~$5-$20
    • The Free Pascal / Lazarus port in ..\lazarus\ is the recommended
      alternative if you do not already own a VB6 license.

SOURCE FILES
------------
  SongManager.vbp    VB6 project file (open this in VB6 IDE)
  frmMain.frm        Main form — UI and all event handlers
  modAPI.bas         Win32 API declarations
  modCakewalk.bas    Cakewalk window detection & command dispatch
  modPlaylist.bas    Queue engine + rolling buffer state machine
  modDB.bas          CSV / Excel database importer
  modSet.bas         Save / Load .SET playlist files

BUILD — VB6 IDE
---------------
  1. Open SongManager.vbp in the VB6 IDE
  2. Ensure COMDLG32.OCX is registered:
       regsvr32 C:\Windows\System32\COMDLG32.OCX
  3. File → Make SongManager.exe

BUILD — COMMAND LINE (silent compile)
--------------------------------------
  "C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe" ^
      /make SongManager.vbp /outdir dist\

  Output: dist\SongManager.exe

BUILD REQUIREMENTS
------------------
  • Visual Basic 6 SP6
  • COMDLG32.OCX registered (included with Win98 / VB6 runtime)
  • For Excel import: MDAC 2.6+ and Microsoft Jet 4.0 OLEDB provider

RUNTIME REQUIREMENTS
--------------------
  • Windows 98 / ME  (Win32 only — no .NET required)
  • Cakewalk Professional 3.0 must be running before clicking Play
  • VB6 runtime: MSVBVM60.DLL (distributed with VB6 applications)

KNOWN LIMITATIONS
-----------------
  • MIDI Transpose (Tr+/Tr-) display works but audio effect is a stub.
  • Multi-select in the Add dialog returns only the first file.
  • Cakewalk WM_COMMAND IDs may need adjustment for some CW 3.0 builds.

======================================================================
