unit uSet;

{$mode objfpc}{$H+}

interface

uses
  SysUtils;

// Write current queue to a .SET file. Returns 0 on success, error code otherwise.
function SaveSet(const sFilePath: string): Integer;

// Read a .SET file and append entries to the queue.
// Does NOT clear queue first — caller clears if needed.
function LoadSet(const sFilePath: string;
                 out   nLoaded:   Integer;
                 out   sError:    string): Boolean;

implementation

uses
  uPlaylist, StrUtils;

const
  SET_SIGNATURE = '# SongManager Set File';
  FIELD_SEP     = '|';

// Simple pipe-delimited splitter
function SplitPipe(const s: string; out f0, f1, f2: string): Integer;
var
  p1, p2: Integer;
begin
  f0 := ''; f1 := ''; f2 := '';
  p1 := Pos('|', s);
  if p1 = 0 then
  begin
    f0 := s;
    Result := 1;
    Exit;
  end;
  f0 := Copy(s, 1, p1 - 1);
  p2 := PosEx('|', s, p1 + 1);
  if p2 = 0 then
  begin
    f1 := Copy(s, p1 + 1, MaxInt);
    Result := 2;
    Exit;
  end;
  f1 := Copy(s, p1 + 1, p2 - p1 - 1);
  f2 := Copy(s, p2 + 1, MaxInt);
  Result := 3;
end;

function SaveSet(const sFilePath: string): Integer;
var
  f: TextFile;
  i: Integer;
begin
  Result := 0;
  AssignFile(f, sFilePath);
  try
    Rewrite(f);
    try
      WriteLn(f, SET_SIGNATURE);
      WriteLn(f, '# Saved: ' + FormatDateTime('YYYY-MM-DD hh:nn', Now));
      WriteLn(f, '# Songs: ' + IntToStr(g_QueueSize));
      WriteLn(f, '');
      for i := 0 to g_QueueSize - 1 do
        WriteLn(f, g_Queue[i].FullPath + FIELD_SEP +
                   g_Queue[i].Artist   + FIELD_SEP +
                   IntToStr(GetDelay(i)));
    finally
      CloseFile(f);
    end;
  except
    on E: Exception do
      Result := 1;
  end;
end;

function LoadSet(const sFilePath: string;
                 out   nLoaded:   Integer;
                 out   sError:    string): Boolean;
var
  f: TextFile;
  sLine: string;
  nLine: Integer;
  sPath, sArtist, sDelayStr: string;
  nDelay, qi: Integer;
begin
  nLoaded := 0;
  sError  := '';
  Result  := False;

  if not FileExists(sFilePath) then
  begin
    sError := 'File not found: ' + sFilePath;
    Exit;
  end;

  AssignFile(f, sFilePath);
  nLine := 0;
  try
    Reset(f);
    try
      while not Eof(f) do
      begin
        ReadLn(f, sLine);
        Inc(nLine);
        sLine := Trim(sLine);

        if (sLine = '') or (sLine[1] = '#') then Continue;

        SplitPipe(sLine, sPath, sArtist, sDelayStr);
        sPath   := Trim(sPath);
        sArtist := Trim(sArtist);
        nDelay  := StrToIntDef(Trim(sDelayStr), 0);

        if (UpperCase(ExtractFileExt(sPath)) <> '.WRK') or
           (Pos('\', sPath) = 0) then Continue;

        qi := AddSong(sPath, sArtist);
        SetDelay(qi, nDelay);
        Inc(nLoaded);
      end;
    finally
      CloseFile(f);
    end;
    Result := True;
  except
    on E: Exception do
      sError := 'Set load error on line ' + IntToStr(nLine) + ': ' + E.Message;
  end;
end;

end.
