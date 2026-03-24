unit uMain;

{$mode objfpc}{$H+}

interface

uses
  Windows,
  Classes, SysUtils,
  Forms, Controls, Graphics, Dialogs,
  StdCtrls, ExtCtrls,
  uCakewalk;  // for TCWState

type
  { TfrmMain }
  TfrmMain = class(TForm)
    // Search bar
    txtSearch     : TEdit;
    lblSearchHint : TLabel;
    // Song list
    lstSongs      : TListBox;
    // Status
    lblStatus     : TLabel;
    // Transport row
    btnPlay       : TButton;
    btnStop       : TButton;
    btnPause      : TButton;
    btnNext       : TButton;
    btnTrMinus    : TButton;
    lblTranspose  : TLabel;
    btnTrPlus     : TButton;
    // Playlist management row
    btnAdd        : TButton;
    btnDelete     : TButton;
    btnDelay      : TButton;
    btnClear      : TButton;
    btnLoadDB     : TButton;
    btnLoadSet    : TButton;
    btnSaveSet    : TButton;
    // Font / path row
    btnFontPlus   : TButton;
    btnFontMinus  : TButton;
    chkFullPath   : TCheckBox;
    // Non-visual
    tmrPoll       : TTimer;
    tmrConnect    : TTimer;
    dlgOpen       : TOpenDialog;
    dlgSave       : TSaveDialog;

    // Events wired from .lfm
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tmrPollTimer(Sender: TObject);
    procedure tmrConnectTimer(Sender: TObject);
    procedure btnPlayClick(Sender: TObject);
    procedure btnStopClick(Sender: TObject);
    procedure btnPauseClick(Sender: TObject);
    procedure btnNextClick(Sender: TObject);
    procedure btnTrPlusClick(Sender: TObject);
    procedure btnTrMinusClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnDeleteClick(Sender: TObject);
    procedure btnDelayClick(Sender: TObject);
    procedure btnClearClick(Sender: TObject);
    procedure btnLoadDBClick(Sender: TObject);
    procedure btnLoadSetClick(Sender: TObject);
    procedure btnSaveSetClick(Sender: TObject);
    procedure btnFontPlusClick(Sender: TObject);
    procedure btnFontMinusClick(Sender: TObject);
    procedure chkFullPathClick(Sender: TObject);
    procedure txtSearchChange(Sender: TObject);
    procedure txtSearchKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure lstSongsDblClick(Sender: TObject);

  private
    m_TransposeSemitones : Integer;
    m_bShowFullPath      : Boolean;
    m_bUpdatingList      : Boolean;
    m_PrevState          : uCakewalk.TCWState;

    procedure UpdateConnectionStatus;
    procedure RefreshList;
    function  SelectedQueueIdx: Integer;
    function  NowPlayingLabel: string;
    function  FormatTranspose(n: Integer): string;
  end;

var
  frmMain: TfrmMain;

implementation

uses
  uPlaylist, uDB, uSet;

// Dark-theme color helpers (Windows COLORREF = 0x00BBGGRR)
function ColRGB(R, G, B: Byte): TColor; inline;
begin
  Result := TColor((B shl 16) or (G shl 8) or R);
end;

const
  CLR_BG        = $00202020; // very dark gray background
  CLR_BG_MID    = $00404040; // medium dark gray (form bg)
  CLR_BG_CTRL   = $00202020; // control background
  CLR_FG_GREEN  = $0000FF00; // green text (playing / normal)
  CLR_FG_GRAY   = $00C0C0C0; // light gray text
  CLR_FG_RED    = $004040FF; // reddish text (error)  BGR: R=FF,G=40,B=40
  CLR_FG_CYAN   = $00FFFF00; // cyan text (paused)    BGR: R=0,G=FF,B=FF

procedure TfrmMain.FormCreate(Sender: TObject);
begin
  m_TransposeSemitones := 0;
  m_bShowFullPath      := False;
  m_bUpdatingList      := False;
  m_PrevState          := uCakewalk.cwUnknown;

  // ---- Form ----
  Color      := CLR_BG_MID;
  Font.Color := CLR_FG_GRAY;

  // ---- Search bar ----
  txtSearch.Color      := CLR_BG_CTRL;
  txtSearch.Font.Color := CLR_FG_GREEN;
  lblSearchHint.Color  := CLR_BG_MID;
  lblSearchHint.Font.Color := CLR_FG_GRAY;

  // ---- Song list ----
  lstSongs.Color      := CLR_BG_CTRL;
  lstSongs.Font.Name  := 'Courier New';
  lstSongs.Font.Size  := 12;
  lstSongs.Font.Color := CLR_FG_GREEN;

  // ---- Status bar ----
  lblStatus.Color      := CLR_BG_CTRL;
  lblStatus.Font.Color := CLR_FG_RED;
  lblStatus.Alignment  := taCenter;
  lblStatus.Caption    := 'Cakewalk: NOT FOUND';

  // ---- Transport buttons ----
  btnPlay.Color  := ColRGB(0,   96,  0);
  btnStop.Color  := ColRGB(96,  0,   0);
  btnPause.Color := ColRGB(80,  80,  0);

  // ---- Transpose label ----
  lblTranspose.Color      := CLR_BG_CTRL;
  lblTranspose.Font.Color := $00FFFF00; // cyan
  lblTranspose.Caption    := ' 0';
  lblTranspose.Alignment  := taCenter;

  // ---- Init playlist engine ----
  InitPlaylist;

  // ---- Timers ----
  tmrPoll.Interval    := 500;
  tmrConnect.Interval := 2000;
  tmrPoll.Enabled     := True;
  tmrConnect.Enabled  := True;

  // ---- Initial connection check ----
  UpdateConnectionStatus;
end;

procedure TfrmMain.FormDestroy(Sender: TObject);
begin
  tmrPoll.Enabled    := False;
  tmrConnect.Enabled := False;
end;

// ---------------------------------------------------------------------------
// Polling timer (500 ms) — drives rolling buffer
// ---------------------------------------------------------------------------
procedure TfrmMain.tmrPollTimer(Sender: TObject);
var
  sPos:     string;
  newState: uCakewalk.TCWState;
begin
  if not IsCakewalkRunning then Exit;

  newState := uCakewalk.PollState(sPos);

  case newState of
    uCakewalk.cwPlaying:
      begin
        lblStatus.Caption    := 'PLAYING  [' + sPos + ']  —  ' + NowPlayingLabel;
        lblStatus.Font.Color := CLR_FG_GREEN;
      end;
    uCakewalk.cwPaused:
      begin
        lblStatus.Caption    := 'PAUSED   [' + sPos + ']';
        lblStatus.Font.Color := CLR_FG_CYAN;
      end;
    uCakewalk.cwStopped:
      begin
        lblStatus.Caption    := 'STOPPED';
        lblStatus.Font.Color := CLR_FG_GRAY;
      end;
  end;

  // Transition Playing → Stopped = song ended
  if (m_PrevState = uCakewalk.cwPlaying) and (newState = uCakewalk.cwStopped) then
  begin
    OnSongAdvanced;
    RefreshList;
  end;

  m_PrevState := newState;
end;

// ---------------------------------------------------------------------------
// Connection timer (2 s)
// ---------------------------------------------------------------------------
procedure TfrmMain.tmrConnectTimer(Sender: TObject);
begin
  UpdateConnectionStatus;
end;

procedure TfrmMain.UpdateConnectionStatus;
begin
  if IsCakewalkRunning then
  begin
    if not (m_PrevState in [uCakewalk.cwPlaying, uCakewalk.cwPaused]) then
    begin
      if g_QueueSize > 0 then
        lblStatus.Caption := 'Cakewalk: CONNECTED  —  ' + IntToStr(g_QueueSize) + ' song(s) in queue'
      else
        lblStatus.Caption := 'Cakewalk: CONNECTED  —  Queue empty';
      lblStatus.Font.Color := CLR_FG_GRAY;
    end;
  end
  else
  begin
    lblStatus.Caption    := 'Cakewalk: NOT FOUND  —  Please start Cakewalk 3.0';
    lblStatus.Font.Color := CLR_FG_RED;
  end;
end;

// ---------------------------------------------------------------------------
// Transport buttons
// ---------------------------------------------------------------------------
procedure TfrmMain.btnPlayClick(Sender: TObject);
var
  startIdx: Integer;
begin
  if not IsCakewalkRunning then
  begin
    MessageDlg('Cakewalk is not running.', mtWarning, [mbOK], 0);
    Exit;
  end;
  if g_QueueSize = 0 then
  begin
    MessageDlg('The queue is empty. Add songs first.', mtInformation, [mbOK], 0);
    Exit;
  end;
  if g_QueueIdx < 0 then
  begin
    startIdx := SelectedQueueIdx;
    if startIdx < 0 then startIdx := 0;
    StartPlayback(startIdx);
  end
  else
    CW_Play;
  RefreshList;
end;

procedure TfrmMain.btnStopClick(Sender: TObject);
begin
  CW_Stop;
  RefreshList;
end;

procedure TfrmMain.btnPauseClick(Sender: TObject);
begin
  CW_Pause;
end;

procedure TfrmMain.btnNextClick(Sender: TObject);
begin
  SkipToNext;
  RefreshList;
end;

// ---------------------------------------------------------------------------
// Transpose stubs (Phase 2)
// ---------------------------------------------------------------------------
procedure TfrmMain.btnTrPlusClick(Sender: TObject);
begin
  if m_TransposeSemitones < 12 then
  begin
    Inc(m_TransposeSemitones);
    lblTranspose.Caption := FormatTranspose(m_TransposeSemitones);
  end;
end;

procedure TfrmMain.btnTrMinusClick(Sender: TObject);
begin
  if m_TransposeSemitones > -12 then
  begin
    Dec(m_TransposeSemitones);
    lblTranspose.Caption := FormatTranspose(m_TransposeSemitones);
  end;
end;

function TfrmMain.FormatTranspose(n: Integer): string;
begin
  if n > 0 then Result := '+' + IntToStr(n)
  else           Result := IntToStr(n);
end;

// ---------------------------------------------------------------------------
// Playlist management buttons
// ---------------------------------------------------------------------------
procedure TfrmMain.btnAddClick(Sender: TObject);
begin
  dlgOpen.Title  := 'Add Cakewalk File';
  dlgOpen.Filter := 'Cakewalk Files (*.wrk)|*.WRK|All Files (*.*)|*.*';
  dlgOpen.FilterIndex := 1;
  if not dlgOpen.Execute then Exit;
  AddSong(dlgOpen.FileName, '');
  RefreshList;
end;

procedure TfrmMain.btnDeleteClick(Sender: TObject);
var
  qi: Integer;
begin
  qi := SelectedQueueIdx;
  if qi < 0 then Exit;
  if qi = g_QueueIdx then
    if MessageDlg('Delete the currently playing song?',
                  mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;
  RemoveSong(qi);
  RefreshList;
end;

procedure TfrmMain.btnDelayClick(Sender: TObject);
var
  qi: Integer;
  sVal: string;
begin
  qi := SelectedQueueIdx;
  if qi < 0 then Exit;
  sVal := IntToStr(GetDelay(qi));
  if InputQuery('Set Delay',
                'Pre-play delay in seconds for:' + LineEnding + g_Queue[qi].DisplayName,
                sVal) then
    SetDelay(qi, StrToIntDef(sVal, 0));
end;

procedure TfrmMain.btnClearClick(Sender: TObject);
begin
  if g_QueueSize > 0 then
    if MessageDlg('Clear the entire queue?',
                  mtConfirmation, [mbYes, mbNo], 0) <> mrYes then Exit;
  ClearQueue;
  RefreshList;
end;

procedure TfrmMain.btnLoadDBClick(Sender: TObject);
var
  nAdded: Integer;
  sError: string;
begin
  dlgOpen.Title  := 'Load Song Database';
  dlgOpen.Filter := 'Database Files (*.csv;*.xls)|*.CSV;*.XLS|CSV Files (*.csv)|*.CSV|Excel Files (*.xls)|*.XLS';
  dlgOpen.FilterIndex := 1;
  if not dlgOpen.Execute then Exit;
  if DetectAndLoad(dlgOpen.FileName, nAdded, sError) then
  begin
    MessageDlg('Loaded ' + IntToStr(nAdded) + ' song(s) from database.',
               mtInformation, [mbOK], 0);
    RefreshList;
  end
  else
    MessageDlg('Import failed:' + LineEnding + sError, mtError, [mbOK], 0);
end;

procedure TfrmMain.btnSaveSetClick(Sender: TObject);
var
  sPath: string;
begin
  dlgSave.Title      := 'Save Set';
  dlgSave.Filter     := 'SongManager Set Files (*.set)|*.SET|All Files (*.*)|*.*';
  dlgSave.DefaultExt := 'set';
  if not dlgSave.Execute then Exit;
  sPath := dlgSave.FileName;
  if UpperCase(ExtractFileExt(sPath)) <> '.SET' then
    sPath := sPath + '.SET';
  if SaveSet(sPath) = 0 then
    MessageDlg('Set saved: ' + sPath, mtInformation, [mbOK], 0)
  else
    MessageDlg('Save failed.', mtError, [mbOK], 0);
end;

procedure TfrmMain.btnLoadSetClick(Sender: TObject);
var
  nLoaded: Integer;
  sError:  string;
begin
  dlgOpen.Title  := 'Load Set';
  dlgOpen.Filter := 'SongManager Set Files (*.set)|*.SET|All Files (*.*)|*.*';
  dlgOpen.FilterIndex := 1;
  if not dlgOpen.Execute then Exit;
  if LoadSet(dlgOpen.FileName, nLoaded, sError) then
  begin
    MessageDlg('Loaded ' + IntToStr(nLoaded) + ' song(s) from set.',
               mtInformation, [mbOK], 0);
    RefreshList;
  end
  else
    MessageDlg('Load failed:' + LineEnding + sError, mtError, [mbOK], 0);
end;

// ---------------------------------------------------------------------------
// Font size controls
// ---------------------------------------------------------------------------
procedure TfrmMain.btnFontPlusClick(Sender: TObject);
begin
  if lstSongs.Font.Size < 36 then
    lstSongs.Font.Size := lstSongs.Font.Size + 2;
end;

procedure TfrmMain.btnFontMinusClick(Sender: TObject);
begin
  if lstSongs.Font.Size > 6 then
    lstSongs.Font.Size := lstSongs.Font.Size - 2;
end;

// ---------------------------------------------------------------------------
// Full-path toggle
// ---------------------------------------------------------------------------
procedure TfrmMain.chkFullPathClick(Sender: TObject);
begin
  m_bShowFullPath := chkFullPath.Checked;
  RefreshList;
end;

// ---------------------------------------------------------------------------
// Search box
// ---------------------------------------------------------------------------
procedure TfrmMain.txtSearchChange(Sender: TObject);
begin
  BuildFilter(txtSearch.Text);
  RefreshList;
end;

procedure TfrmMain.txtSearchKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  qi: Integer;
begin
  case Key of
    VK_RETURN:
      begin
        qi := SelectedQueueIdx;
        if (qi >= 0) and IsCakewalkRunning then
        begin
          StartPlayback(qi);
          RefreshList;
          txtSearch.Text := '';
          BuildFilter('');
        end;
      end;
    VK_DOWN:
      if lstSongs.ItemIndex < lstSongs.Count - 1 then
        lstSongs.ItemIndex := lstSongs.ItemIndex + 1;
    VK_UP:
      if lstSongs.ItemIndex > 0 then
        lstSongs.ItemIndex := lstSongs.ItemIndex - 1;
    VK_ESCAPE:
      begin
        txtSearch.Text := '';
        BuildFilter('');
        RefreshList;
      end;
  end;
end;

// ---------------------------------------------------------------------------
// List double-click = start playback
// ---------------------------------------------------------------------------
procedure TfrmMain.lstSongsDblClick(Sender: TObject);
var
  qi: Integer;
begin
  qi := SelectedQueueIdx;
  if qi < 0 then Exit;
  if not IsCakewalkRunning then
  begin
    MessageDlg('Cakewalk is not running.', mtWarning, [mbOK], 0);
    Exit;
  end;
  StartPlayback(qi);
  RefreshList;
end;

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
procedure TfrmMain.RefreshList;
var
  i, qi, count: Integer;
  prevQIdx: Integer;
  mappedQ: Integer;
begin
  if m_bUpdatingList then Exit;
  m_bUpdatingList := True;
  try
    prevQIdx := SelectedQueueIdx;

    lstSongs.Items.BeginUpdate;
    lstSongs.Items.Clear;

    if g_FilterActive then count := g_FilteredSize
    else                   count := g_QueueSize;

    for i := 0 to count - 1 do
    begin
      if g_FilterActive then qi := g_FilteredIdx[i]
      else                   qi := i;
      lstSongs.Items.Add(GetDisplayLabel(qi, m_bShowFullPath));
    end;
    lstSongs.Items.EndUpdate;

    // Restore selection
    if prevQIdx >= 0 then
      for i := 0 to lstSongs.Count - 1 do
      begin
        mappedQ := FilteredToQueueIdx(i);
        if mappedQ = prevQIdx then
        begin
          lstSongs.ItemIndex := i;
          Break;
        end;
      end;

    // Fall back to PLAYING row
    if (lstSongs.ItemIndex < 0) and (g_QueueIdx >= 0) then
      for i := 0 to lstSongs.Count - 1 do
        if FilteredToQueueIdx(i) = g_QueueIdx then
        begin
          lstSongs.ItemIndex := i;
          Break;
        end;
  finally
    m_bUpdatingList := False;
  end;
end;

function TfrmMain.SelectedQueueIdx: Integer;
begin
  if lstSongs.ItemIndex < 0 then
    Result := -1
  else
    Result := FilteredToQueueIdx(lstSongs.ItemIndex);
end;

function TfrmMain.NowPlayingLabel: string;
begin
  if (g_QueueIdx >= 0) and (g_QueueIdx < g_QueueSize) then
  begin
    Result := g_Queue[g_QueueIdx].DisplayName;
    if g_Queue[g_QueueIdx].Artist <> '' then
      Result := Result + '  (' + g_Queue[g_QueueIdx].Artist + ')';
  end
  else
    Result := '';
end;

end.
