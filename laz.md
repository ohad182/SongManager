older structure
C:\Users\ohadc\SongManager\
├── vb\                    ← original VB6 (untouched)
│   ├── SongManager.vbp
│   ├── frmMain.frm
│   ├── modAPI.bas
│   ├── modCakewalk.bas
│   ├── modPlaylist.bas
│   ├── modDB.bas
│   ├── modSet.bas
│   └── README.txt
└── lazarus\               ← new Free Pascal port
    ├── SongManager.lpi    (project file)
    ├── SongManager.lpr    (program entry)
    ├── uAPI.pas
    ├── uCakewalk.pas
    ├── uPlaylist.pas
    ├── uDB.pas
    ├── uSet.pas
    ├── uMain.pas
    └── uMain.lfm
How to build
Option A — Lazarus IDE (easiest, recommended)
Install Lazarus (free, ~250 MB)
Open lazarus\SongManager.lpi
Press F9 (Run) or Shift+F9 (Build only)
Output: lazarus\SongManager.exe
Option B — lazbuild command line (after Lazarus is installed)
Open a PowerShell or cmd window and run:

cd C:\Users\ohadc\SongManager\lazarus
& "C:\lazarus\lazbuild.exe" SongManager.lpi
That's it. lazbuild handles the .res file generation, finds all units, and produces SongManager.exe in the same folder.

Option C — release build (optimized, no debug info)
& "C:\lazarus\lazbuild.exe" --build-mode=Release SongManager.lpi
One-time install steps before building
Download Lazarus installer from lazarus-ide.org/index.php?page=downloads — pick the 64-bit Windows bundle (includes FPC compiler)
Run the installer — default path is C:\lazarus
No registration, no license key
After that, the lazbuild command above will produce a working SongManager.exe targeting modern Windows (10/11), keeping all the same Cakewalk integration logic from the original VB6 code.