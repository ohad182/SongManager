program SongManager;

{$mode objfpc}{$H+}

uses
  Interfaces,  // LCL widgetset (Win32 on Windows)
  Forms,
  uMain    { TfrmMain },
  uAPI     { Win32 helpers },
  uCakewalk{ Cakewalk integration },
  uPlaylist{ Queue engine },
  uDB      { CSV/XLS import },
  uSet     { .SET file I/O };

{$R *.res}

begin
  RequireDerivedFormResource := True;
  Application.Initialize;
  Application.MainFormOnTaskBar := True;
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
