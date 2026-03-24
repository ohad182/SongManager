Attribute VB_Name = "modAPI"
'===========================================================================
' modAPI.bas  —  Win32 API declarations for SongManager
' Target: Windows 98 / ME  (VB6 SP6)
'===========================================================================
Option Explicit

' ── Window Management ───────────────────────────────────────────────────────
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
    (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, _
     ByVal lpszClass As String, ByVal lpszWindow As String) As Long

Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
    (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" _
    (ByVal hWnd As Long) As Long

Public Declare Function SetForegroundWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function IsWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function ShowWindow Lib "user32" _
    (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long

Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, lpRect As RECT) As Long

' ── Child Window Enumeration ─────────────────────────────────────────────────
Public Declare Function EnumChildWindows Lib "user32" _
    (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, _
     ByVal lParam As Long) As Long

' ── Message Sending ──────────────────────────────────────────────────────────
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SendMessageStr Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As String) As Long

Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long

' ── Keyboard Injection ───────────────────────────────────────────────────────
Public Declare Sub keybd_event Lib "user32" _
    (ByVal bVk As Byte, ByVal bScan As Byte, _
     ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

' ── Clipboard (used to pass long filenames to Cakewalk safely) ───────────────
Public Declare Function OpenClipboard Lib "user32" _
    (ByVal hWnd As Long) As Long

Public Declare Function CloseClipboard Lib "user32" () As Long

Public Declare Function EmptyClipboard Lib "user32" () As Long

Public Declare Function SetClipboardData Lib "user32" _
    (ByVal wFormat As Long, ByVal hMem As Long) As Long

Public Declare Function GlobalAlloc Lib "kernel32" _
    (ByVal wFlags As Long, ByVal dwBytes As Long) As Long

Public Declare Function GlobalLock Lib "kernel32" _
    (ByVal hMem As Long) As Long

Public Declare Function GlobalUnlock Lib "kernel32" _
    (ByVal hMem As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)

' ── Delay without freezing UI ────────────────────────────────────────────────
Public Declare Sub Sleep Lib "kernel32" _
    (ByVal dwMilliseconds As Long)

' ── File/Path helpers ────────────────────────────────────────────────────────
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, _
     ByVal cchBuffer As Long) As Long

' ── Structures ───────────────────────────────────────────────────────────────
Public Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type

' ── Windows Message Constants ────────────────────────────────────────────────
Public Const WM_COMMAND    As Long = &H111
Public Const WM_CLOSE      As Long = &H10
Public Const WM_SETTEXT    As Long = &HC
Public Const WM_GETTEXT    As Long = &HD
Public Const WM_KEYDOWN    As Long = &H100
Public Const WM_KEYUP      As Long = &H101
Public Const WM_CHAR       As Long = &H102
Public Const WM_LBUTTONDOWN As Long = &H201
Public Const WM_LBUTTONUP   As Long = &H202

' ── ShowWindow constants ─────────────────────────────────────────────────────
Public Const SW_SHOW       As Long = 5
Public Const SW_RESTORE    As Long = 9

' ── Virtual Key Codes ────────────────────────────────────────────────────────
Public Const VK_SPACE      As Byte = &H20
Public Const VK_RETURN     As Byte = &HD
Public Const VK_ESCAPE     As Byte = &H1B
Public Const VK_CONTROL    As Byte = &H11
Public Const VK_SHIFT      As Byte = &H10
Public Const VK_F5         As Byte = &H74
Public Const VK_DELETE     As Byte = &H2E
Public Const VK_A          As Byte = &H41
Public Const VK_O          As Byte = &H4F
Public Const VK_P          As Byte = &H50
Public Const VK_S          As Byte = &H53

' ── keybd_event flags ────────────────────────────────────────────────────────
Public Const KEYEVENTF_KEYDOWN  As Long = &H0
Public Const KEYEVENTF_KEYUP    As Long = &H2

' ── Clipboard format ─────────────────────────────────────────────────────────
Public Const CF_TEXT       As Long = 1
Public Const GMEM_MOVEABLE As Long = &H2
Public Const GMEM_ZEROINIT As Long = &H40

' ── Helper: read text from any window handle ─────────────────────────────────
Public Function GetWndText(ByVal hWnd As Long) As String
    Dim nLen  As Long
    Dim sText As String
    nLen = GetWindowTextLength(hWnd) + 1
    If nLen <= 1 Then
        GetWndText = ""
        Exit Function
    End If
    sText = String(nLen, Chr(0))
    GetWindowText hWnd, sText, nLen
    GetWndText = Left(sText, nLen - 1)
End Function

' ── Helper: send Ctrl+key to a window ────────────────────────────────────────
Public Sub SendCtrlKey(ByVal hWnd As Long, ByVal vk As Byte)
    SetForegroundWindow hWnd
    Sleep 50
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYDOWN, 0
    keybd_event vk, 0, KEYEVENTF_KEYDOWN, 0
    Sleep 30
    keybd_event vk, 0, KEYEVENTF_KEYUP, 0
    keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
    Sleep 50
End Sub

' ── Helper: press a single key (no modifier) ─────────────────────────────────
Public Sub SendKey(ByVal hWnd As Long, ByVal vk As Byte)
    SetForegroundWindow hWnd
    Sleep 50
    keybd_event vk, 0, KEYEVENTF_KEYDOWN, 0
    Sleep 30
    keybd_event vk, 0, KEYEVENTF_KEYUP, 0
    Sleep 50
End Sub

' ── Helper: paste text to the focused edit control via clipboard ──────────────
' Used to feed long filenames into Cakewalk's Open/Add dialogs
Public Sub PasteTextToClipboard(ByVal sText As String)
    Dim hMem    As Long
    Dim hLock   As Long
    Dim abData() As Byte
    Dim nLen    As Long

    abData = StrConv(sText & Chr(0), vbFromUnicode)
    nLen = UBound(abData) + 1

    hMem = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, nLen)
    If hMem = 0 Then Exit Sub

    hLock = GlobalLock(hMem)
    CopyMemory ByVal hLock, abData(0), nLen
    GlobalUnlock hMem

    OpenClipboard 0
    EmptyClipboard
    SetClipboardData CF_TEXT, hMem
    CloseClipboard
End Sub

' ── Helper: convert a long path to its DOS 8.3 equivalent ────────────────────
Public Function ToShortPath(ByVal sLong As String) As String
    Dim sShort As String
    Dim nRet   As Long
    sShort = String(260, Chr(0))
    nRet = GetShortPathName(sLong, sShort, 260)
    If nRet > 0 Then
        ToShortPath = Left(sShort, nRet)
    Else
        ToShortPath = sLong
    End If
End Function
