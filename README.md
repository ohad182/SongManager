======================================================================
SongManager v1.0  —  Cakewalk 3.0 Playlist Wrapper
======================================================================

OVERVIEW
--------
SongManager bypasses Cakewalk 3.0's 128-song limit and 8-character
filename restriction by managing an unlimited external queue and
maintaining a rolling 2-song buffer inside Cakewalk's native playlist.

REPOSITORY LAYOUT
-----------------
  vb\        Original Visual Basic 6 source (Windows 98 / ME target)
  lazarus\   Free Pascal / Lazarus port (Windows 10/11, free toolchain)

WHICH VERSION SHOULD I USE?
----------------------------
  Lazarus    Recommended. Free toolchain, builds on modern Windows,
             same functionality, actively maintainable.

  VB6        Only if you already have Visual Basic 6 SP6 installed and
             need to target Windows 98 / ME specifically.

BUILD — LAZARUS (FREE)
-----------------------
  1. Download Lazarus from https://lazarus-ide.org  (64-bit Windows
     installer, ~250 MB, includes the Free Pascal compiler)
  2. Install (default path: C:\lazarus)

  IDE build:
    Open  lazarus\SongManager.lpi  in Lazarus IDE
    Press F9  (or Shift+F9 for build-only)
    Output: lazarus\SongManager.exe

  Command-line build:
    cd C:\Users\ohadc\SongManager\lazarus
    C:\lazarus\lazbuild.exe SongManager.lpi

  Release build (optimised):
    C:\lazarus\lazbuild.exe --build-mode=Release SongManager.lpi

BUILD — VB6 (LEGACY)
---------------------
  See vb\README.txt for full instructions.

  Requires Visual Basic 6 SP6 (commercial, no longer sold by Microsoft).
  Short form:
    "C:\Program Files\Microsoft Visual Studio\VB98\VB6.exe" /make vb\SongManager.vbp

======================================================================
