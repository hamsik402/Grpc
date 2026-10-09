unit UnitMain;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Grids, Data.DB, Data.Win.ADODB, UnitBon, UnitTiers;

type
  TFrmMain = class(TForm)
    pnlTop: TPanel;
    lblTitle: TLabel;
    pnlActions: TPanel;
    btnRefresh: TButton;
    btnAdd: TButton;
    btnEntry: TButton;
    btnExit: TButton;
    btnLivraison: TButton;
    btnAffectation: TButton;
    btnFournisseurs: TButton;
    btnServices: TButton;
    edtCode: TEdit;
    edtDesignation: TEdit;
    edtType: TEdit;
    edtQty: TEdit;
    lblCode: TLabel;
    lblDesignation: TLabel;
    lblType: TLabel;
    lblQty: TLabel;
    grdArticles: TStringGrid;
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    procedure FormCreate(Sender: TObject);
    procedure btnRefreshClick(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnEntryClick(Sender: TObject);
    procedure btnExitClick(Sender: TObject);
    procedure btnLivraisonClick(Sender: TObject);
    procedure btnAffectationClick(Sender: TObject);
    procedure btnFournisseursClick(Sender: TObject);
    procedure btnServicesClick(Sender: TObject);
  private
    procedure ConnectDatabase;
    procedure LoadArticles;
    function ReadQty: Integer;
  public
  end;

var
  FrmMain: TFrmMain;

implementation

{$R *.dfm}

procedure TFrmMain.ConnectDatabase;
var
  DBPath: string;
begin
  DBPath := ExtractFilePath(ParamStr(0)) + 'StockNaftal.accdb';
  ADOConnection1.LoginPrompt := False;
  ADOConnection1.ConnectionString :=
    'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=' + DBPath +
    ';Persist Security Info=False;';
  ADOConnection1.Connected := True;
  ADOQuery1.Connection := ADOConnection1;
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
  R: Integer;
begin
  if not ADOConnection1.Connected then Exit;
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'SELECT Code_art, Designation, Type_art, Qtt_stock, Seuil_alerte ' +
    'FROM ARTICLE ORDER BY Designation';
  ADOQuery1.Open;
  grdArticles.RowCount := ADOQuery1.RecordCount + 1;
  R := 1;
  while not ADOQuery1.Eof do
  begin
    grdArticles.Cells[0,R] := ADOQuery1.FieldByName('Code_art').AsString;
    grdArticles.Cells[1,R] := ADOQuery1.FieldByName('Designation').AsString;
    grdArticles.Cells[2,R] := ADOQuery1.FieldByName('Type_art').AsString;
    grdArticles.Cells[3,R] := ADOQuery1.FieldByName('Qtt_stock').AsString;
    grdArticles.Cells[4,R] := ADOQuery1.FieldByName('Seuil_alerte').AsString;
    Inc(R);
    ADOQuery1.Next;
  end;
end;

function TFrmMain.ReadQty: Integer;
begin
  if not TryStrToInt(Trim(edtQty.Text), Result) or (Result <= 0) then
    raise Exception.Create('La quantité doit être un entier supérieur à zéro.');
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

procedure TFrmMain.btnEntryClick(Sender: TObject);
var
  Q: Integer;
  Code: string;
begin
  Q := ReadQty;
  Code := Trim(edtCode.Text);
  if Code = '' then raise Exception.Create('Saisissez le code de l''article.');
  ADOQuery1.Close;
  ADOQuery1.SQL.Text :=
    'UPDATE ARTICLE SET Qtt_stock = Qtt_stock + :q WHERE Code_art = :c';
  ADOQuery1.Parameters.ParamByName('q').Value := Q;
  ADOQuery1.Parameters.ParamByName('c').Value := Code;
  ADOQuery1.ExecSQL;
  if ADOQuery1.RowsAffected = 0 then
    raise Exception.Create('Code article introuvable.');
  LoadArticles;
  ShowMessage('Entrée de stock enregistrée.');
end;

procedure TFrmMain.btnExitClick(Sender: TObject);
var
  Q: Integer;
  Code: string;
  CurrentQty: Integer;
begin
  Q := ReadQty;
  Code := Trim(edtCode.Text);
  if Code = '' then raise Exception.Create('Saisissez le code de l''article.');

  ADOQuery1.Close;
  ADOQuery1.SQL.Text := 'SELECT Qtt_stock FROM ARTICLE WHERE Code_art = :c';
  ADOQuery1.Parameters.ParamByName('c').Value := Code;
  ADOQuery1.Open;
  if ADOQuery1.Eof then
    raise Exception.Create('Code article introuvable.');
  CurrentQty := ADOQuery1.FieldByName('Qtt_stock').AsInteger;
  ADOQuery1.Close;

  if Q > CurrentQty then
    raise Exception.CreateFmt('Stock insuffisant. Disponible : %d', [CurrentQty]);

  ADOQuery1.SQL.Text :=
    'UPDATE ARTICLE SET Qtt_stock = Qtt_stock - :q WHERE Code_art = :c';
  ADOQuery1.Parameters.ParamByName('q').Value := Q;
  ADOQuery1.Parameters.ParamByName('c').Value := Code;
  ADOQuery1.ExecSQL;
  LoadArticles;
  ShowMessage('Sortie de stock enregistrée.');
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

end.
