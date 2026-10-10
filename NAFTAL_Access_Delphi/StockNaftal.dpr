program StockNaftal;

uses
  Vcl.Forms,
  UnitMain in 'UnitMain.pas' {FrmMain},
  UnitBon in 'UnitBon.pas',
  UnitTiers in 'UnitTiers.pas',
  UnitConsult in 'UnitConsult.pas',
  UnitStock in 'UnitStock.pas',
  UnitMouv in 'UnitMouv.pas',
  UnitRappro in 'UnitRappro.pas',
  UnitSchema in 'UnitSchema.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.Title := 'Gestion de stock - NAFTAL Bouira';
  Application.CreateForm(TFrmMain, FrmMain);
  Application.Run;
end.
