unit uAPI;

{$mode objfpc}{$H+}

interface

uses
  Windows, SysUtils;

// Helper: read text from any window handle
function GetWndText(hWnd: HWND): string;

// Helper: send Ctrl+key to a window
procedure SendCtrlKey(hWnd: HWND; vk: Byte);

// Helper: press a single key (no modifier)
procedure SendKey(hWnd: HWND; vk: Byte);

// Helper: paste text to clipboard (used to feed long filenames into Cakewalk dialogs)
procedure PasteTextToClipboard(const sText: string);

// Helper: convert a long path to its DOS 8.3 equivalent
function ToShortPath(const sLong: string): string;

implementation

function GetWndText(hWnd: HWND): string;
var
  nLen: Integer;
  buf: array[0..1023] of AnsiChar;
begin
  nLen := GetWindowTextLength(hWnd);
  if nLen = 0 then
  begin
    Result := '';
    Exit;
  end;
  FillChar(buf, SizeOf(buf), 0);
  GetWindowText(hWnd, buf, SizeOf(buf) - 1);
  Result := string(buf);
end;

procedure SendCtrlKey(hWnd: HWND; vk: Byte);
begin
  SetForegroundWindow(hWnd);
  Sleep(50);
  keybd_event(VK_CONTROL, 0, 0, 0);
  keybd_event(vk, 0, 0, 0);
  Sleep(30);
  keybd_event(vk, 0, KEYEVENTF_KEYUP, 0);
  keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0);
  Sleep(50);
end;

procedure SendKey(hWnd: HWND; vk: Byte);
begin
  SetForegroundWindow(hWnd);
  Sleep(50);
  keybd_event(vk, 0, 0, 0);
  Sleep(30);
  keybd_event(vk, 0, KEYEVENTF_KEYUP, 0);
  Sleep(50);
end;

procedure PasteTextToClipboard(const sText: string);
var
  hMem: THandle;
  pMem: PAnsiChar;
  sAnsi: AnsiString;
  nLen: Integer;
begin
  sAnsi := AnsiString(sText);
  nLen  := Length(sAnsi) + 1;
  hMem  := GlobalAlloc(GMEM_MOVEABLE or GMEM_ZEROINIT, nLen);
  if hMem = 0 then Exit;
  pMem := GlobalLock(hMem);
  if pMem <> nil then
  begin
    Move(sAnsi[1], pMem^, nLen);
    GlobalUnlock(hMem);
  end;
  OpenClipboard(0);
  EmptyClipboard;
  SetClipboardData(CF_TEXT, hMem);
  CloseClipboard;
end;

function ToShortPath(const sLong: string): string;
var
  buf: array[0..MAX_PATH] of AnsiChar;
  nRet: DWORD;
begin
  nRet := GetShortPathName(PAnsiChar(AnsiString(sLong)), buf, MAX_PATH);
  if nRet > 0 then
    Result := string(buf)
  else
    Result := sLong;
end;

end.
