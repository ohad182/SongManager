unit uDB;

{$mode objfpc}{$H+}

interface

uses
  SysUtils;

// Detect file format by extension and load into queue.
// Returns True on success; nAdded = songs added; sError filled on failure.
function DetectAndLoad(const sFilePath: string;
                       out   nAdded:    Integer;
                       out   sError:    string): Boolean;

function LoadFromCSV(const sFilePath: string;
                     out   nAdded:    Integer;
                     out   sError:    string): Boolean;

function LoadFromExcel(const sFilePath: string;
                       out   nAdded:    Integer;
                       out   sError:    string): Boolean;

implementation

uses
  uPlaylist, ComObj, Variants;

// --------------------------------------------------------------------------
// CSV field splitter — handles quoted fields
// --------------------------------------------------------------------------
procedure SplitCSVLine(const sLine: string; var cols: array of string;
                       out nCols: Integer);
var
  i: Integer;
  ch: Char;
  bQuote: Boolean;
  sCurrent: string;
const
  MAX_COLS = 20;
begin
  nCols    := 0;
  bQuote   := False;
  sCurrent := '';

  for i := 1 to Length(sLine) do
  begin
    ch := sLine[i];
    if bQuote then
    begin
      if ch = '"' then
        bQuote := False
      else
        sCurrent := sCurrent + ch;
    end
    else
    begin
      if ch = '"' then
        bQuote := True
      else if ch = ',' then
      begin
        if nCols <= MAX_COLS then
        begin
          cols[nCols] := sCurrent;
          Inc(nCols);
        end;
        sCurrent := '';
      end
      else
        sCurrent := sCurrent + ch;
    end;
  end;
  // last field
  if nCols <= MAX_COLS then
  begin
    cols[nCols] := sCurrent;
    Inc(nCols);
  end;
end;

// --------------------------------------------------------------------------
// CSV loader
// --------------------------------------------------------------------------
function LoadFromCSV(const sFilePath: string;
                     out   nAdded:    Integer;
                     out   sError:    string): Boolean;
var
  f: TextFile;
  sLine: string;
  nLine: Integer;
  cols: array[0..20] of string;
  nCols: Integer;
  sName, sArtist, sPath: string;
begin
  nAdded := 0;
  sError := '';
  Result := False;

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

        // skip blanks and comments
        if (sLine = '') or (sLine[1] = '#') then Continue;

        SplitCSVLine(sLine, cols, nCols);
        if nCols < 3 then Continue;

        sName   := Trim(cols[0]);
        sArtist := Trim(cols[1]);
        sPath   := Trim(cols[2]);

        // Auto-detect header row: no backslash in the path column
        if (nLine = 1) and (Pos('\', sPath) = 0) then Continue;

        // Must be a .WRK file with a path
        if (UpperCase(ExtractFileExt(sPath)) <> '.WRK') or
           (Pos('\', sPath) = 0) then Continue;

        AddSong(sPath, sArtist);
        Inc(nAdded);
      end;
    finally
      CloseFile(f);
    end;
    Result := True;
  except
    on E: Exception do
    begin
      sError := 'CSV read error on line ' + IntToStr(nLine) + ': ' + E.Message;
    end;
  end;
end;

// --------------------------------------------------------------------------
// Excel loader (ADO + Jet 4.0, requires MDAC 2.6+)
// --------------------------------------------------------------------------
function GetFirstSheetName(conn: OleVariant): string;
var
  cat: OleVariant;
  i: Integer;
  tblName: string;
begin
  Result := 'Sheet1$'; // safe default
  try
    cat := CreateOleObject('ADOX.Catalog');
    cat.ActiveConnection := conn;
    for i := 0 to cat.Tables.Count - 1 do
    begin
      tblName := VarToStr(cat.Tables[i].Name);
      if (Length(tblName) > 0) and (tblName[Length(tblName)] = '$') then
      begin
        Result := tblName;
        Exit;
      end;
    end;
  except
    // fall back to default
  end;
end;

function LoadFromExcel(const sFilePath: string;
                       out   nAdded:    Integer;
                       out   sError:    string): Boolean;
var
  conn, rs: OleVariant;
  sSheet: string;
  j, iName, iArtist, iPath: Integer;
  sFieldUp, sPath, sArtist: string;
begin
  nAdded := 0;
  sError := '';
  Result := False;

  if not FileExists(sFilePath) then
  begin
    sError := 'File not found: ' + sFilePath;
    Exit;
  end;

  try
    conn := CreateOleObject('ADODB.Connection');
    conn.Open('Provider=Microsoft.Jet.OLEDB.4.0;' +
              'Data Source=' + sFilePath + ';' +
              'Extended Properties="Excel 8.0;HDR=Yes;IMEX=1;"');

    sSheet := GetFirstSheetName(conn);

    rs := CreateOleObject('ADODB.Recordset');
    rs.Open('SELECT * FROM [' + sSheet + ']', conn, 0, 1);

    // Map column names to indices
    iName := -1; iArtist := -1; iPath := -1;
    for j := 0 to rs.Fields.Count - 1 do
    begin
      sFieldUp := UpperCase(Trim(VarToStr(rs.Fields[j].Name)));
      if (sFieldUp = 'SONGNAME') or (sFieldUp = 'SONG NAME') or
         (sFieldUp = 'SONG')     or (sFieldUp = 'NAME') or
         (sFieldUp = 'TITLE') then
        iName := j
      else if (sFieldUp = 'ARTIST') or (sFieldUp = 'PERFORMER') or
              (sFieldUp = 'BAND') then
        iArtist := j
      else if (sFieldUp = 'FILEPATH') or (sFieldUp = 'FILE PATH') or
              (sFieldUp = 'PATH')     or (sFieldUp = 'FILENAME') or
              (sFieldUp = 'FILE') then
        iPath := j;
    end;

    // Positional fallback
    if (iName   = -1) and (rs.Fields.Count >= 3) then iName   := 0;
    if (iArtist = -1) and (rs.Fields.Count >= 3) then iArtist := 1;
    if (iPath   = -1) and (rs.Fields.Count >= 3) then iPath   := 2;

    if iPath = -1 then
    begin
      sError := 'Cannot find FilePath column in sheet: ' + sSheet;
      rs.Close;
      conn.Close;
      Exit;
    end;

    while not rs.EOF do
    begin
      sPath := Trim(VarToStr(rs.Fields[iPath].Value));
      if iArtist >= 0 then
        sArtist := Trim(VarToStr(rs.Fields[iArtist].Value))
      else
        sArtist := '';

      if (UpperCase(ExtractFileExt(sPath)) = '.WRK') and
         (Pos('\', sPath) > 0) then
      begin
        AddSong(sPath, sArtist);
        Inc(nAdded);
      end;
      rs.MoveNext;
    end;

    rs.Close;
    conn.Close;
    Result := True;
  except
    on E: Exception do
      sError := 'Excel import error: ' + E.Message;
  end;
end;

// --------------------------------------------------------------------------
// Auto-detect and dispatch
// --------------------------------------------------------------------------
function DetectAndLoad(const sFilePath: string;
                       out   nAdded:    Integer;
                       out   sError:    string): Boolean;
var
  sExt: string;
begin
  sExt := UpperCase(ExtractFileExt(Trim(sFilePath)));
  case sExt of
    '.XLS': Result := LoadFromExcel(sFilePath, nAdded, sError);
    else    Result := LoadFromCSV(sFilePath, nAdded, sError);
  end;
end;

end.
