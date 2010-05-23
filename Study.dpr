program Study;

uses
  Forms,
  fMain in 'fMain.pas' {frmMain},
  fAbout in 'fAbout.pas' {frmAbout};

{$R *.res}
{$R XpManifest.res}

begin
  Application.Initialize;
  Application.CreateForm(TfrmMain, frmMain);
  Application.CreateForm(TfrmAbout, frmAbout);
  Application.Run;
end.
