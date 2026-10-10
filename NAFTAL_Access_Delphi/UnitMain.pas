unit UnitMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Grids, Data.DB, Data.Win.ADODB, UnitBon, UnitTiers, UnitConsult, UnitStock, UnitMouv, UnitRappro, UnitSchema;

type
  TFrmMain = class(TForm)
    pnlTop: TPanel;
    lblTitle: TLabel;
    pnlActions: TPanel;
    btnRefresh: TButton;
    btnAdd: TButton;
    btnLivraison: TButton;
    btnAffectation: TButton;
    btnFournisseurs: TButton;
    btnServices: TButton;
    btnConsLiv: TButton;
    btnConsAff: TButton;
    btnCommande: TButton;
    btnConsCmd: TButton;
    btnStock: TButton;
    btnMouv: TButton;
    btnRappro: TButton;
    lblAlert: TLabel;
    edtCode: TEdit;
    edtDesignation: TEdit;
    edtType: TEdit;
    lblCode: TLabel;
    lblDesignation: TLabel;
    lblType: TLabel;
    grdArticles: TStringGrid;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnLivraisonClick(Sender: TObject);
    procedure btnAffectationClick(Sender: TObject);
    procedure btnFournisseursClick(Sender: TObject);
    procedure btnServicesClick(Sender: TObject);
    procedure btnConsLivClick(Sender: TObject);
    procedure btnConsAffClick(Sender: TObject);
    procedure btnCommandeClick(Sender: TObject);
    procedure btnConsCmdClick(Sender: TObject);
    procedure btnStockClick(Sender: TObject);
    procedure btnMouvClick(Sender: TObject);
    procedure btnRapproClick(Sender: TObject);
  private
    procedure ConnectDatabase;
    procedure LoadArticles;
  public
  end;

var
  FrmMain: TFrmMain;

implementation

{$R *.dfm}

{ Ouvre la base ; la crée au premier lancement et complète les tables ou
  colonnes manquantes (voir UnitSchema). Essaie ACE 12.0 puis 16.0. }
procedure TFrmMain.ConnectDatabase;
var
  DBPath, Prov, LastErr: string;
  Created: Boolean;
begin
  DBPath := EnsureDatabase(Created);
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.Connected := False;
  LastErr := '';
  for Prov in ProviderList do
  begin
    try
      ADOConnection1.ConnectionString :=
        'Provider=' + Prov + ';Data Source=' + DBPath +
        ';Persist Security Info=False;';
      ADOConnection1.Connected := True;
      Break;
    except
      on E: Exception do
        LastErr := E.Message;
    end;
  end;
  if not ADOConnection1.Connected then
    raise Exception.Create(LastErr);
  ADOQuery1.Connection := ADOConnection1;
  EnsureSchema(ADOConnection1);
  if Created then
    ShowMessage('Première utilisation : la base de données a été créée ici :' +
      sLineBreak + DBPath);
end;

procedure TFrmMain.FormCreate(Sender: TObject);
begin
  Caption := 'Gestion de stock - NAFTAL Bouira';
  grdArticles.ColCount := 5;
  grdArticles.RowCount := 2;
  grdArticles.FixedRows := 1;
  grdArticles.Cells[0,0] := 'Code';
  grdArticles.Cells[1,0] := 'Désignation';
  grdArticles.Cells[2,0] := 'Type';
  grdArticles.Cells[3,0] := 'Quantité';
  grdArticles.Cells[4,0] := 'Seuil';
  try
    ConnectDatabase;
    LoadArticles;
  except
    on E: Exception do
      MessageDlg('Connexion impossible : ' + E.Message + sLineBreak +
        'Vérifiez StockNaftal.accdb et le fournisseur ACE.', mtError, [mbOK], 0);
  end;
end;

procedure TFrmMain.LoadArticles;
var
  R, Alerts, Qtt: Integer;
begin
  if not ADOConnection1.Connected then Exit;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT Code_art, Designation, Type_art, Qtt_stock, Seuil_alerte ' +
    'FROM ARTICLE ORDER BY Designation';
  ADOQuery1.Open;
  if ADOQuery1.RecordCount < 1 then
  begin
    grdArticles.RowCount := 2;
    for R := 0 to grdArticles.ColCount - 1 do
      grdArticles.Cells[R, 1] := '';
  end
  else
    grdArticles.RowCount := ADOQuery1.RecordCount + 1;
  Alerts := 0;
  R := 1;
  while not ADOQuery1.Eof do
  begin
    grdArticles.Cells[0,R] := ADOQuery1.FieldByName('Code_art').AsString;
    grdArticles.Cells[1,R] := ADOQuery1.FieldByName('Designation').AsString;
    grdArticles.Cells[2,R] := ADOQuery1.FieldByName('Type_art').AsString;
    grdArticles.Cells[3,R] := ADOQuery1.FieldByName('Qtt_stock').AsString;
    grdArticles.Cells[4,R] := ADOQuery1.FieldByName('Seuil_alerte').AsString;
    Qtt := ADOQuery1.FieldByName('Qtt_stock').AsInteger;
    if (Qtt <= 0) or (not ADOQuery1.FieldByName('Seuil_alerte').IsNull and
       (Qtt <= ADOQuery1.FieldByName('Seuil_alerte').AsInteger)) then
      Inc(Alerts);
    Inc(R);
    ADOQuery1.Next;
  end;
  if Alerts > 0 then
    lblAlert.Caption := Format('Alerte : %d article(s) sous le seuil ou en rupture', [Alerts])
  else
    lblAlert.Caption := '';
end;

procedure TFrmMain.btnRefreshClick(Sender: TObject);
begin
  LoadArticles;
end;

procedure TFrmMain.btnAddClick(Sender: TObject);
begin
  if Trim(edtCode.Text) = '' then
    raise Exception.Create('Saisissez le code article.');
  if Trim(edtDesignation.Text) = '' then
    raise Exception.Create('Saisissez la désignation.');
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');

  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'INSERT INTO ARTICLE (Code_art, Designation, Type_art, Qtt_stock, Seuil_alerte) ' +
    'VALUES (:c, :d, :t, 0, 5)';
  ADOQuery1.Parameters.ParamByName('c').Value := Trim(edtCode.Text);
  ADOQuery1.Parameters.ParamByName('d').Value := Trim(edtDesignation.Text);
  ADOQuery1.Parameters.ParamByName('t').Value := Trim(edtType.Text);
  ADOQuery1.ExecSQL;
  LoadArticles;
  ShowMessage('Article ajouté.');
end;

procedure TFrmMain.btnLivraisonClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenBon(ADOConnection1, bkLivraison);
  LoadArticles;
end;

procedure TFrmMain.btnAffectationClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenBon(ADOConnection1, bkAffectation);
  LoadArticles;
end;

procedure TFrmMain.btnFournisseursClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenTiers(ADOConnection1, tkFournisseur);
end;

procedure TFrmMain.btnServicesClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenTiers(ADOConnection1, tkService);
end;

procedure TFrmMain.btnConsLivClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenConsult(ADOConnection1, bkLivraison);
  LoadArticles;
end;

procedure TFrmMain.btnConsAffClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenConsult(ADOConnection1, bkAffectation);
  LoadArticles;
end;

procedure TFrmMain.btnCommandeClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenBon(ADOConnection1, bkCommande);
end;

procedure TFrmMain.btnConsCmdClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenConsult(ADOConnection1, bkCommande);
  LoadArticles;
end;

procedure TFrmMain.btnStockClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenStock(ADOConnection1);
  LoadArticles;
end;

procedure TFrmMain.btnMouvClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenMouvements(ADOConnection1, Trim(edtCode.Text));
end;

procedure TFrmMain.btnRapproClick(Sender: TObject);
begin
  if not ADOConnection1.Connected then
    raise Exception.Create('Base de données non connectée.');
  OpenRappro(ADOConnection1);
end;

end.
