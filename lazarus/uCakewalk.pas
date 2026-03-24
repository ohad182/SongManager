unit uCakewalk;

{$mode objfpc}{$H+}

interface

uses
  Windows, uAPI;

const
  CW_CLASS_NAME    = 'CakeWalk';
  CW_CMD_PLAY      = 32880;
  CW_CMD_STOP      = 32881;
  CW_CMD_PAUSE     = 32882;
  CW_CMD_REWIND    = 32883;
  CW_CMD_OPEN_FILE = 57601;
  CW_CMD_PLAYLIST  = 32920;

type
  TCWState = (cwUnknown, cwStopped, cwPlaying, cwPaused);

// Connection
function  FindCakewalk: Boolean;
function  IsCakewalkRunning: Boolean;
function  CakewalkHandle: HWND;

// Polling
function  PollState(out sPosition: string): TCWState;
function  LastState: TCWState;

// Transport
procedure CW_Play;
procedure CW_Stop;
procedure CW_Pause;
procedure CW_Rewind;

// File loading
procedure CW_OpenFile(const sPath: string);
procedure CW_AddToPlaylist(const sPath: string);
procedure CW_ClearPlaylist;
procedure CW_OpenPlaylistWindow;

// EnumChildWindows callbacks — must be at unit level for FPC stdcall
function EnumTransportProc(hWnd: HWND; lParam: LPARAM): BOOL; stdcall;
function EnumTopLevelProc(hWnd: HWND; lParam: LPARAM): BOOL; stdcall;

// Globals used by callbacks
var
  g_PositionText  : string;
  g_SearchPartial : string;
  g_SearchResult  : HWND;

implementation

var
  m_hCakewalk    : HWND     = 0;
  m_LastTitle    : string   = '';
  m_LastState    : TCWState = cwUnknown;
  m_LastPosition : string   = '';

function CakewalkHandle: HWND;
begin
  Result := m_hCakewalk;
end;

function LastState: TCWState;
begin
  Result := m_LastState;
end;

// --------------------------------------------------------------------------
// Callback: scan child windows for a bar:beat:tick position string
// --------------------------------------------------------------------------
function EnumTransportProc(hWnd: HWND; lParam: LPARAM): BOOL; stdcall;
var
  s: string;
  i, colonCount: Integer;
begin
  s := GetWndText(hWnd);
  if Length(s) >= 7 then
  begin
    colonCount := 0;
    for i := 1 to Length(s) do
      if s[i] = ':' then Inc(colonCount);
    if colonCount >= 2 then
    begin
      g_PositionText := Trim(s);
      Result := False; // stop enumeration
      Exit;
    end;
  end;
  Result := True; // continue
end;

// --------------------------------------------------------------------------
// Callback: find window whose title contains g_SearchPartial
// --------------------------------------------------------------------------
function EnumTopLevelProc(hWnd: HWND; lParam: LPARAM): BOOL; stdcall;
var
  s: string;
begin
  s := GetWndText(hWnd);
  if Pos(UpperCase(g_SearchPartial), UpperCase(s)) > 0 then
  begin
    g_SearchResult := hWnd;
    Result := False; // stop
    Exit;
  end;
  Result := True; // continue
end;

// --------------------------------------------------------------------------
// Private helpers
// --------------------------------------------------------------------------
function ReadTransportPosition: string;
begin
  g_PositionText := '';
  EnumChildWindows(m_hCakewalk, @EnumTransportProc, 0);
  Result := g_PositionText;
end;

function FindWindowByPartialTitle(const sPartial: string): HWND;
begin
  g_SearchPartial := sPartial;
  g_SearchResult  := 0;
  EnumWindows(@EnumTopLevelProc, 0);
  Result := g_SearchResult;
end;

function FindWindowByPartialTitleChild(hParent: HWND; const sPartial: string): HWND;
begin
  g_SearchPartial := sPartial;
  g_SearchResult  := 0;
  EnumChildWindows(hParent, @EnumTopLevelProc, 0);
  Result := g_SearchResult;
end;

function FindPlaylistWindow: HWND;
var
  hMDI: HWND;
begin
  hMDI := FindWindowEx(m_hCakewalk, 0, 'MDIClient', nil);
  if hMDI = 0 then
  begin
    Result := FindWindowEx(m_hCakewalk, 0, nil, 'Playlist');
    Exit;
  end;
  Result := FindWindowEx(hMDI, 0, nil, 'Playlist');
  if Result = 0 then
    Result := FindWindowByPartialTitleChild(hMDI, 'Playlist');
end;

function FindOpenDialog: HWND;
begin
  Result := FindWindow('#32770', nil);
  if Result <> 0 then Exit;
  Result := FindWindowEx(m_hCakewalk, 0, '#32770', nil);
end;

// --------------------------------------------------------------------------
// Public: connection
// --------------------------------------------------------------------------
function FindCakewalk: Boolean;
begin
  m_hCakewalk := FindWindow(CW_CLASS_NAME, nil);
  if m_hCakewalk = 0 then
    m_hCakewalk := FindWindowByPartialTitle('Cakewalk');
  Result := m_hCakewalk <> 0;
end;

function IsCakewalkRunning: Boolean;
begin
  if m_hCakewalk = 0 then
  begin
    Result := FindCakewalk;
    Exit;
  end;
  if not Boolean(IsWindow(m_hCakewalk)) then
  begin
    m_hCakewalk := 0;
    Result := False;
  end
  else
    Result := True;
end;

// --------------------------------------------------------------------------
// Public: polling
// --------------------------------------------------------------------------
function PollState(out sPosition: string): TCWState;
var
  sTitle, sUpper: string;
  newState: TCWState;
begin
  if not IsCakewalkRunning then
  begin
    Result := cwUnknown;
    Exit;
  end;

  sTitle    := GetWndText(m_hCakewalk);
  sPosition := ReadTransportPosition;
  sUpper    := UpperCase(sTitle);

  if Pos('PLAYING', sUpper) > 0 then
    newState := cwPlaying
  else if (Pos('PAUSED', sUpper) > 0) or (Pos('PAUSE', sUpper) > 0) then
    newState := cwPaused
  else if (Pos('STOPPED', sUpper) > 0) or (Pos('STOP', sUpper) > 0) then
    newState := cwStopped
  else
    newState := m_LastState;

  // Position-reset detection: 1:01 at start of a new song
  if (m_LastState = cwPlaying) and (newState = cwPlaying) then
    if (Copy(sPosition, 1, 4) = '1:01') and (Copy(m_LastPosition, 1, 4) <> '1:01') then
      newState := cwStopped;

  m_LastTitle    := sTitle;
  m_LastState    := newState;
  m_LastPosition := sPosition;
  Result         := newState;
end;

// --------------------------------------------------------------------------
// Public: transport
// --------------------------------------------------------------------------
procedure CW_Play;
begin
  if not IsCakewalkRunning then Exit;
  SendMessage(m_hCakewalk, WM_COMMAND, CW_CMD_PLAY, 0);
end;

procedure CW_Stop;
begin
  if not IsCakewalkRunning then Exit;
  SendMessage(m_hCakewalk, WM_COMMAND, CW_CMD_STOP, 0);
  SendKey(m_hCakewalk, VK_SPACE);
end;

procedure CW_Pause;
begin
  if not IsCakewalkRunning then Exit;
  SendMessage(m_hCakewalk, WM_COMMAND, CW_CMD_PAUSE, 0);
end;

procedure CW_Rewind;
begin
  if not IsCakewalkRunning then Exit;
  SendMessage(m_hCakewalk, WM_COMMAND, CW_CMD_REWIND, 0);
end;

// --------------------------------------------------------------------------
// Public: file loading
// --------------------------------------------------------------------------
procedure CW_OpenFile(const sPath: string);
var
  hDlg, hEdit: HWND;
begin
  if not IsCakewalkRunning then Exit;

  ShowWindow(m_hCakewalk, SW_RESTORE);
  SetForegroundWindow(m_hCakewalk);
  Sleep(150);

  PasteTextToClipboard(sPath);

  SendCtrlKey(m_hCakewalk, Ord('O'));
  Sleep(400);

  hDlg := FindOpenDialog;
  if hDlg = 0 then Exit;

  hEdit := FindWindowEx(hDlg, 0, 'Edit', nil);
  if hEdit = 0 then Exit;

  SendMessage(hEdit, WM_SETTEXT, 0, 0);
  Sleep(30);
  SetForegroundWindow(hDlg);
  Sleep(30);
  SendCtrlKey(hDlg, Ord('A'));
  Sleep(30);
  SendCtrlKey(hDlg, Ord('V'));
  Sleep(80);
  SendKey(hDlg, VK_RETURN);
  Sleep(300);
end;

procedure CW_AddToPlaylist(const sPath: string);
var
  hPL, hAddDlg, hEdit: HWND;
begin
  if not IsCakewalkRunning then Exit;

  CW_OpenPlaylistWindow;
  Sleep(200);

  hPL := FindPlaylistWindow;
  if hPL = 0 then Exit;

  PasteTextToClipboard(sPath);

  SetForegroundWindow(hPL);
  Sleep(80);
  SendMessage(hPL, WM_COMMAND, 1001, 0);
  Sleep(200);

  hAddDlg := FindOpenDialog;
  if hAddDlg <> 0 then
  begin
    hEdit := FindWindowEx(hAddDlg, 0, 'Edit', nil);
    if hEdit <> 0 then
    begin
      SendMessage(hEdit, WM_SETTEXT, 0, 0);
      Sleep(30);
      SetForegroundWindow(hAddDlg);
      SendCtrlKey(hAddDlg, Ord('V'));
      Sleep(80);
      SendKey(hAddDlg, VK_RETURN);
      Sleep(300);
    end;
  end;
end;

procedure CW_ClearPlaylist;
var
  hPL: HWND;
begin
  if not IsCakewalkRunning then Exit;
  CW_OpenPlaylistWindow;
  Sleep(200);
  hPL := FindPlaylistWindow;
  if hPL = 0 then Exit;
  SetForegroundWindow(hPL);
  Sleep(80);
  SendCtrlKey(hPL, Ord('A'));
  Sleep(80);
  SendKey(hPL, VK_DELETE);
  Sleep(100);
end;

procedure CW_OpenPlaylistWindow;
var
  hPL: HWND;
begin
  hPL := FindPlaylistWindow;
  if (hPL <> 0) and Boolean(IsWindowVisible(hPL)) then Exit;
  SendMessage(m_hCakewalk, WM_COMMAND, CW_CMD_PLAYLIST, 0);
  Sleep(150);
end;

end.
