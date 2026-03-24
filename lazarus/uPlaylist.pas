unit uPlaylist;

{$mode objfpc}{$H+}

interface

uses
  SysUtils;

type
  TSongEntry = record
    DisplayName : string;
    FullPath    : string;
    Artist      : string;
    Delay       : Integer;  // pre-play pause in seconds
  end;

var
  g_Queue        : array of TSongEntry;
  g_QueueSize    : Integer;
  g_QueueIdx     : Integer;   // index of currently PLAYING song; -1 = none
  g_FilteredIdx  : array of Integer;
  g_FilteredSize : Integer;
  g_FilterActive : Boolean;

// Lifecycle
procedure InitPlaylist;
procedure ClearQueue;

// Queue mutation
function  AddSong(const sFullPath, sArtist: string): Integer;
procedure RemoveSong(queueIndex: Integer);
procedure SetDelay(queueIndex: Integer; nSeconds: Integer);
function  GetDelay(queueIndex: Integer): Integer;

// Playback control (delegates to uCakewalk)
function  StartPlayback(queueIndex: Integer): Boolean;
procedure OnSongAdvanced;
procedure SkipToNext;

// Query helpers
function  IsPlayingIndex(queueIndex: Integer): Boolean;
function  IsReadyIndex(queueIndex: Integer): Boolean;
function  GetDisplayLabel(queueIndex: Integer; bShowFullPath: Boolean): string;

// Search / filter
procedure BuildFilter(const sSearch: string);
function  FilteredToQueueIdx(filteredPos: Integer): Integer;

// Utility
function StripToDisplayName(const sPath: string): string;

implementation

uses
  uCakewalk, Windows;

procedure InitPlaylist;
begin
  g_QueueSize    := 0;
  g_QueueIdx     := -1;
  g_FilterActive := False;
  g_FilteredSize := 0;
  SetLength(g_Queue, 0);
  SetLength(g_FilteredIdx, 0);
end;

procedure ClearQueue;
begin
  InitPlaylist;
end;

function StripToDisplayName(const sPath: string): string;
begin
  Result := ChangeFileExt(ExtractFileName(sPath), '');
  Result := StringReplace(Result, '_', ' ', [rfReplaceAll]);
end;

function AddSong(const sFullPath, sArtist: string): Integer;
begin
  SetLength(g_Queue, g_QueueSize + 1);
  g_Queue[g_QueueSize].FullPath    := sFullPath;
  g_Queue[g_QueueSize].Artist      := sArtist;
  g_Queue[g_QueueSize].DisplayName := StripToDisplayName(sFullPath);
  g_Queue[g_QueueSize].Delay       := 0;
  Result      := g_QueueSize;
  g_QueueSize := g_QueueSize + 1;
end;

procedure RemoveSong(queueIndex: Integer);
var
  i: Integer;
begin
  if (queueIndex < 0) or (queueIndex >= g_QueueSize) then Exit;

  for i := queueIndex to g_QueueSize - 2 do
    g_Queue[i] := g_Queue[i + 1];

  Dec(g_QueueSize);
  SetLength(g_Queue, g_QueueSize);

  if queueIndex < g_QueueIdx then
    Dec(g_QueueIdx)
  else if queueIndex = g_QueueIdx then
  begin
    if g_QueueIdx >= g_QueueSize then
      g_QueueIdx := g_QueueSize - 1;
  end;
end;

procedure SetDelay(queueIndex: Integer; nSeconds: Integer);
begin
  if (queueIndex >= 0) and (queueIndex < g_QueueSize) then
    g_Queue[queueIndex].Delay := nSeconds;
end;

function GetDelay(queueIndex: Integer): Integer;
begin
  if (queueIndex >= 0) and (queueIndex < g_QueueSize) then
    Result := g_Queue[queueIndex].Delay
  else
    Result := 0;
end;

function StartPlayback(queueIndex: Integer): Boolean;
begin
  if g_QueueSize = 0 then
  begin
    Result := False;
    Exit;
  end;
  if queueIndex < 0 then queueIndex := 0;
  if queueIndex >= g_QueueSize then
  begin
    Result := False;
    Exit;
  end;

  g_QueueIdx := queueIndex;

  CW_ClearPlaylist;
  Sleep(200);

  CW_AddToPlaylist(g_Queue[g_QueueIdx].FullPath);
  if g_QueueIdx + 1 < g_QueueSize then
    CW_AddToPlaylist(g_Queue[g_QueueIdx + 1].FullPath);

  Sleep(150);

  // Apply pre-play delay if set
  if g_Queue[g_QueueIdx].Delay > 0 then
    Sleep(g_Queue[g_QueueIdx].Delay * 1000);

  CW_Play;
  Result := True;
end;

procedure OnSongAdvanced;
var
  nextIdx: Integer;
begin
  if g_QueueIdx < 0 then Exit;
  Inc(g_QueueIdx);
  nextIdx := g_QueueIdx + 1;
  if nextIdx < g_QueueSize then
    CW_AddToPlaylist(g_Queue[nextIdx].FullPath);
end;

procedure SkipToNext;
begin
  if g_QueueIdx + 1 >= g_QueueSize then Exit;
  CW_Stop;
  Sleep(150);
  StartPlayback(g_QueueIdx + 1);
end;

function IsPlayingIndex(queueIndex: Integer): Boolean;
begin
  Result := queueIndex = g_QueueIdx;
end;

function IsReadyIndex(queueIndex: Integer): Boolean;
begin
  Result := (g_QueueIdx >= 0) and (queueIndex = g_QueueIdx + 1);
end;

function GetDisplayLabel(queueIndex: Integer; bShowFullPath: Boolean): string;
var
  sName: string;
  nPos: Integer;
begin
  if (queueIndex < 0) or (queueIndex >= g_QueueSize) then
  begin
    Result := '';
    Exit;
  end;

  if bShowFullPath then
    sName := g_Queue[queueIndex].FullPath
  else
    sName := g_Queue[queueIndex].DisplayName;

  if IsPlayingIndex(queueIndex) then
    Result := Chr(16) + ' PLAYING  ' + sName
  else if IsReadyIndex(queueIndex) then
    Result := '  READY   ' + sName
  else
  begin
    nPos := queueIndex - g_QueueIdx;
    if nPos < 0 then nPos := queueIndex + 1;
    Result := Format('  %3d       %s', [nPos, sName]);
  end;
end;

procedure BuildFilter(const sSearch: string);
var
  sUp: string;
  i: Integer;
begin
  if Trim(sSearch) = '' then
  begin
    g_FilterActive := False;
    g_FilteredSize := 0;
    Exit;
  end;

  g_FilterActive := True;
  sUp := UpperCase(sSearch);

  SetLength(g_FilteredIdx, g_QueueSize);
  g_FilteredSize := 0;

  for i := 0 to g_QueueSize - 1 do
  begin
    if (Pos(sUp, UpperCase(g_Queue[i].DisplayName)) > 0) or
       (Pos(sUp, UpperCase(g_Queue[i].Artist))      > 0) or
       (Pos(sUp, UpperCase(g_Queue[i].FullPath))    > 0) then
    begin
      g_FilteredIdx[g_FilteredSize] := i;
      Inc(g_FilteredSize);
    end;
  end;
end;

function FilteredToQueueIdx(filteredPos: Integer): Integer;
begin
  if not g_FilterActive then
    Result := filteredPos
  else if (filteredPos >= 0) and (filteredPos < g_FilteredSize) then
    Result := g_FilteredIdx[filteredPos]
  else
    Result := -1;
end;

end.
