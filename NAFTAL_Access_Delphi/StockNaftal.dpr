program StockNaftal;

uses
  Vcl.Forms,
  UnitMain in 'UnitMain.pas' {FrmMain},
  UnitBon in 'UnitBon.pas',
  UnitTiers in 'UnitTiers.pas',
  UnitConsult in 'UnitConsult.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFrmMain, FrmMain);
  Application.Run;
end.
