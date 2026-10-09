unit UnitBon;

{ Bons de livraison (entrée de stock, fournisseur -> magasin) et
  bons d'affectation (sortie de stock, magasin -> service).
  Une seule fiche sert aux deux types de bon ; elle est construite en code.
  L'enregistrement d'un bon (en-tête + lignes + mise à jour du stock) se fait
  dans une seule transaction : tout est enregistré, ou rien. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.DateUtils,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.ComCtrls, Vcl.Grids, Vcl.Dialogs,
  Data.DB, Data.Win.ADODB;

type
  TBonKind = (bkLivraison, bkAffectation, bkCommande);

  TFrmBon = class(TForm)
  private
    FConn: TADOConnection;
    FKind: TBonKind;
    FQry: TADOQuery;
    FTiersCodes: TStringList;
    FArtCodes: TStringList;
    FArtStock: array of Integer;
    FLineCount: Integer;

    edtCode: TEdit;
    dtpDate: TDateTimePicker;
    cboTiers: TComboBox;
    cboArticle: TComboBox;
    edtQty: TEdit;
    edtPrix: TEdit;
    lblStock: TLabel;
    lblTotal: TLabel;
    grdLines: TStringGrid;
    btnAddLine: TButton;
    btnDelLine: TButton;
    btnSave: TButton;
    btnClose: TButton;

    function AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
    function AddEdit(ALeft, ATop, AWidth: Integer): TEdit;
    function AddButton(const ACaption: string; ALeft, ATop, AWidth: Integer;
      AHandler: TNotifyEvent): TButton;
    procedure BuildUI;

    procedure LoadTiers;
    procedure LoadArticles;
    function NextCode: string;
    procedure NewBon;

    function ParseQty(const S: string): Integer;
    function ParsePrix(const S: string): Currency;
    procedure ClearGridRow(ARow: Integer);
    procedure UpdateTotal;
    procedure UpdateStockLabel;

    procedure FormShowHandler(Sender: TObject);
    procedure ArticleChange(Sender: TObject);
    procedure AddLineClick(Sender: TObject);
    procedure DelLineClick(Sender: TObject);
    procedure SaveClick(Sender: TObject);
    procedure CloseClick(Sender: TObject);

    procedure ExecLivraison(const ACode: string);
    procedure ExecAffectation(const ACode: string);
    procedure ExecCommande(const ACode: string);
  public
    constructor CreateBon(AOwner: TComponent; AConn: TADOConnection;
      AKind: TBonKind);
    destructor Destroy; override;
  end;

procedure OpenBon(AConn: TADOConnection; AKind: TBonKind);

implementation

procedure OpenBon(AConn: TADOConnection; AKind: TBonKind);
var
  F: TFrmBon;
begin
  F := TFrmBon.CreateBon(Application, AConn, AKind);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

{ TFrmBon }

constructor TFrmBon.CreateBon(AOwner: TComponent; AConn: TADOConnection;
  AKind: TBonKind);
begin
  inherited CreateNew(AOwner);
  FConn := AConn;
  FKind := AKind;
  FTiersCodes := TStringList.Create;
  FArtCodes := TStringList.Create;
  FQry := TADOQuery.Create(Self);
  FQry.Connection := FConn;
  BuildUI;
  OnShow := FormShowHandler;
  LoadTiers;
  LoadArticles;
  NewBon;
end;

destructor TFrmBon.Destroy;
begin
  FTiersCodes.Free;
  FArtCodes.Free;
  inherited;
end;

function TFrmBon.AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
begin
  Result := TLabel.Create(Self);
  Result.Parent := Self;
  Result.Left := ALeft;
  Result.Top := ATop;
  Result.Caption := ACaption;
end;

function TFrmBon.AddEdit(ALeft, ATop, AWidth: Integer): TEdit;
begin
  Result := TEdit.Create(Self);
  Result.Parent := Self;
  Result.SetBounds(ALeft, ATop, AWidth, 21);
end;

function TFrmBon.AddButton(const ACaption: string; ALeft, ATop, AWidth: Integer;
  AHandler: TNotifyEvent): TButton;
begin
  Result := TButton.Create(Self);
  Result.Parent := Self;
  Result.SetBounds(ALeft, ATop, AWidth, 28);
  Result.Caption := ACaption;
  Result.OnClick := AHandler;
end;

procedure TFrmBon.BuildUI;
var
  IsLiv: Boolean;
begin
  IsLiv := FKind = bkLivraison;
  BorderStyle := bsDialog;
  Position := poOwnerFormCenter;
  ClientWidth := 750;
  ClientHeight := 540;
  case FKind of
    bkLivraison: Caption := 'Bon de livraison (entrée de stock)';
    bkAffectation: Caption := 'Bon d''affectation (sortie de stock)';
  else
    Caption := 'Bon de commande (aucun effet sur le stock)';
  end;

  { En-tête du bon }
  case FKind of
    bkLivraison: AddLabel('N° bon de livraison', 16, 14);
    bkAffectation: AddLabel('N° bon d''affectation', 16, 14);
  else
    AddLabel('N° bon de commande', 16, 14);
  end;
  edtCode := AddEdit(16, 34, 150);
  edtCode.ReadOnly := True;

  AddLabel('Date', 190, 14);
  dtpDate := TDateTimePicker.Create(Self);
  dtpDate.Parent := Self;
  dtpDate.SetBounds(190, 34, 130, 21);
  dtpDate.Date := Date;

  case FKind of
    bkLivraison: AddLabel('Fournisseur', 340, 14);
    bkAffectation: AddLabel('Service destinataire', 340, 14);
  else
    AddLabel('Service demandeur', 340, 14);
  end;
  cboTiers := TComboBox.Create(Self);
  cboTiers.Parent := Self;
  cboTiers.Style := csDropDownList;
  cboTiers.SetBounds(340, 34, 394, 21);

  { Saisie d'une ligne }
  AddLabel('Article', 16, 74);
  cboArticle := TComboBox.Create(Self);
  cboArticle.Parent := Self;
  cboArticle.Style := csDropDownList;
  cboArticle.SetBounds(16, 94, 400, 21);
  cboArticle.OnChange := ArticleChange;

  AddLabel('Quantité', 430, 74);
  edtQty := AddEdit(430, 94, 70);

  if IsLiv then
  begin
    AddLabel('Prix unitaire', 515, 74);
    edtPrix := AddEdit(515, 94, 90);
  end;

  btnAddLine := AddButton('Ajouter la ligne', 624, 90, 110, AddLineClick);

  lblStock := AddLabel('', 16, 122);

  { Lignes du bon }
  grdLines := TStringGrid.Create(Self);
  grdLines.Parent := Self;
  grdLines.SetBounds(16, 146, 718, 290);
  grdLines.FixedCols := 0;
  grdLines.FixedRows := 1;
  grdLines.RowCount := 2;
  grdLines.Options := grdLines.Options + [goRowSelect] - [goEditing];
  if IsLiv then
  begin
    grdLines.ColCount := 4;
    grdLines.ColWidths[0] := 110;
    grdLines.ColWidths[1] := 340;
    grdLines.ColWidths[2] := 80;
    grdLines.ColWidths[3] := 170;
    grdLines.Cells[3, 0] := 'Prix unitaire';
  end
  else
  begin
    grdLines.ColCount := 3;
    grdLines.ColWidths[0] := 120;
    grdLines.ColWidths[1] := 480;
    grdLines.ColWidths[2] := 90;
  end;
  grdLines.Cells[0, 0] := 'Code';
  grdLines.Cells[1, 0] := 'Désignation';
  grdLines.Cells[2, 0] := 'Quantité';

  btnDelLine := AddButton('Supprimer la ligne', 16, 446, 140, DelLineClick);
  lblTotal := AddLabel('', 300, 453);

  btnSave := AddButton('Enregistrer le bon', 480, 496, 140, SaveClick);
  btnClose := AddButton('Fermer', 634, 496, 100, CloseClick);
end;

procedure TFrmBon.FormShowHandler(Sender: TObject);
begin
  if cboTiers.Items.Count = 0 then
  begin
    if FKind = bkLivraison then
      MessageDlg('Aucun fournisseur enregistré. Ajoutez des lignes dans la ' +
        'table FOURNISSEUR (voir 02_donnees_test.sql).', mtWarning, [mbOK], 0)
    else
      MessageDlg('Aucun service enregistré. Ajoutez des lignes dans la ' +
        'table SERVICE_NAFTAL (voir 02_donnees_test.sql).', mtWarning, [mbOK], 0);
  end
  else if cboArticle.Items.Count = 0 then
    MessageDlg('Aucun article enregistré. Ajoutez-en depuis l''écran principal.',
      mtWarning, [mbOK], 0);
end;

procedure TFrmBon.LoadTiers;
var
  Txt: string;
begin
  cboTiers.Items.Clear;
  FTiersCodes.Clear;
  FQry.Close;
  if FKind = bkLivraison then
    FQry.SQL.Text := 'SELECT Code_fourn AS Code, Nom, Prenom FROM FOURNISSEUR ORDER BY Nom'
  else
    FQry.SQL.Text := 'SELECT Code_serv AS Code, Nom_serv AS Nom FROM SERVICE_NAFTAL ORDER BY Nom_serv';
  FQry.Open;
  while not FQry.Eof do
  begin
    Txt := FQry.FieldByName('Code').AsString + ' - ' + FQry.FieldByName('Nom').AsString;
    if FKind = bkLivraison then
      Txt := Trim(Txt + ' ' + FQry.FieldByName('Prenom').AsString);
    cboTiers.Items.Add(Txt);
    FTiersCodes.Add(FQry.FieldByName('Code').AsString);
    FQry.Next;
  end;
  FQry.Close;
end;

procedure TFrmBon.LoadArticles;
var
  Txt: string;
  Stock: Integer;
begin
  cboArticle.Items.Clear;
  FArtCodes.Clear;
  SetLength(FArtStock, 0);
  FQry.Close;
  FQry.SQL.Text :=
    'SELECT Code_art, Designation, Qtt_stock FROM ARTICLE ORDER BY Designation';
  FQry.Open;
  while not FQry.Eof do
  begin
    Stock := FQry.FieldByName('Qtt_stock').AsInteger;
    Txt := FQry.FieldByName('Code_art').AsString + ' - ' +
      FQry.FieldByName('Designation').AsString;
    cboArticle.Items.Add(Txt);
    FArtCodes.Add(FQry.FieldByName('Code_art').AsString);
    SetLength(FArtStock, Length(FArtStock) + 1);
    FArtStock[High(FArtStock)] := Stock;
    FQry.Next;
  end;
  FQry.Close;
  UpdateStockLabel;
end;

function TFrmBon.NextCode: string;
var
  Prefix, Tbl, Fld, S: string;
  N, MaxN: Integer;
begin
  case FKind of
    bkLivraison:
      begin
        Prefix := 'BL-';
        Tbl := 'BON_DE_LIVRAISON';
        Fld := 'Code_Bliv';
      end;
    bkAffectation:
      begin
        Prefix := 'AF-';
        Tbl := 'BON_D_AFFECTATION';
        Fld := 'Code_Aff';
      end;
  else
    begin
      Prefix := 'BC-';
      Tbl := 'BON_DE_COMMANDE';
      Fld := 'Code_BC';
    end;
  end;
  Prefix := Prefix + IntToStr(YearOf(Date)) + '-';

  FQry.Close;
  FQry.SQL.Text := 'SELECT ' + Fld + ' FROM ' + Tbl + ' WHERE ' + Fld + ' LIKE :p';
  FQry.Parameters.ParamByName('p').Value := Prefix + '%';
  FQry.Open;
  MaxN := 0;
  while not FQry.Eof do
  begin
    S := Copy(FQry.Fields[0].AsString, Length(Prefix) + 1, 10);
    if TryStrToInt(S, N) and (N > MaxN) then
      MaxN := N;
    FQry.Next;
  end;
  FQry.Close;
  Result := Prefix + Format('%.4d', [MaxN + 1]);
end;

procedure TFrmBon.NewBon;
var
  R: Integer;
begin
  edtCode.Text := NextCode;
  dtpDate.Date := Date;
  cboTiers.ItemIndex := -1;
  cboArticle.ItemIndex := -1;
  edtQty.Text := '';
  if Assigned(edtPrix) then
    edtPrix.Text := '';
  FLineCount := 0;
  for R := 1 to grdLines.RowCount - 1 do
    ClearGridRow(R);
  grdLines.RowCount := 2;
  UpdateTotal;
  UpdateStockLabel;
end;

function TFrmBon.ParseQty(const S: string): Integer;
begin
  if not TryStrToInt(Trim(S), Result) or (Result <= 0) then
    raise Exception.Create('La quantité doit être un entier supérieur à zéro.');
end;

function TFrmBon.ParsePrix(const S: string): Currency;
var
  T: string;
begin
  { accepte "1250.50" comme "1250,50" ; pas de séparateur de milliers }
  T := StringReplace(Trim(S), ' ', '', [rfReplaceAll]);
  T := StringReplace(T, '.', FormatSettings.DecimalSeparator, [rfReplaceAll]);
  T := StringReplace(T, ',', FormatSettings.DecimalSeparator, [rfReplaceAll]);
  if not TryStrToCurr(T, Result) or (Result < 0) then
    raise Exception.Create('Le prix unitaire doit être un nombre positif ou nul.');
end;

procedure TFrmBon.ClearGridRow(ARow: Integer);
var
  C: Integer;
begin
  for C := 0 to grdLines.ColCount - 1 do
    grdLines.Cells[C, ARow] := '';
end;

procedure TFrmBon.UpdateTotal;
var
  R, Q: Integer;
  SumQty: Integer;
  SumVal: Currency;
begin
  SumQty := 0;
  SumVal := 0;
  for R := 1 to FLineCount do
  begin
    Q := StrToInt(grdLines.Cells[2, R]);
    Inc(SumQty, Q);
    if FKind = bkLivraison then
      SumVal := SumVal + Q * StrToCurr(grdLines.Cells[3, R]);
  end;
  if FKind = bkLivraison then
    lblTotal.Caption := Format('Lignes : %d   Quantité totale : %d   Montant : %s',
      [FLineCount, SumQty, CurrToStrF(SumVal, ffNumber, 2)])
  else
    lblTotal.Caption := Format('Lignes : %d   Quantité totale : %d',
      [FLineCount, SumQty]);
end;

procedure TFrmBon.UpdateStockLabel;
begin
  if cboArticle.ItemIndex >= 0 then
    lblStock.Caption := Format('Stock actuel de cet article : %d',
      [FArtStock[cboArticle.ItemIndex]])
  else
    lblStock.Caption := '';
end;

procedure TFrmBon.ArticleChange(Sender: TObject);
begin
  UpdateStockLabel;
end;

procedure TFrmBon.AddLineClick(Sender: TObject);
var
  Idx, Q, R: Integer;
  Prix: Currency;
  Code: string;
begin
  Idx := cboArticle.ItemIndex;
  if Idx < 0 then
    raise Exception.Create('Choisissez un article.');
  Q := ParseQty(edtQty.Text);
  Prix := 0;
  if FKind = bkLivraison then
    Prix := ParsePrix(edtPrix.Text);

  Code := FArtCodes[Idx];
  for R := 1 to FLineCount do
    if SameText(grdLines.Cells[0, R], Code) then
      raise Exception.Create('Cet article est déjà dans le bon. ' +
        'Supprimez sa ligne pour la ressaisir.');
  if (FKind = bkAffectation) and (Q > FArtStock[Idx]) then
    raise Exception.CreateFmt('Stock insuffisant. Disponible : %d', [FArtStock[Idx]]);

  Inc(FLineCount);
  if grdLines.RowCount < FLineCount + 1 then
    grdLines.RowCount := FLineCount + 1;
  R := FLineCount;
  grdLines.Cells[0, R] := Code;
  grdLines.Cells[1, R] := Copy(cboArticle.Items[Idx],
    Length(Code) + 4, MaxInt);
  grdLines.Cells[2, R] := IntToStr(Q);
  if FKind = bkLivraison then
    grdLines.Cells[3, R] := CurrToStr(Prix);

  edtQty.Text := '';
  if Assigned(edtPrix) then
    edtPrix.Text := '';
  cboArticle.ItemIndex := -1;
  UpdateStockLabel;
  UpdateTotal;
  cboArticle.SetFocus;
end;

procedure TFrmBon.DelLineClick(Sender: TObject);
var
  R, C: Integer;
begin
  R := grdLines.Row;
  if (FLineCount = 0) or (R < 1) or (R > FLineCount) then
    raise Exception.Create('Sélectionnez la ligne à supprimer.');
  for C := R to FLineCount - 1 do
  begin
    grdLines.Rows[C].Assign(grdLines.Rows[C + 1]);
  end;
  ClearGridRow(FLineCount);
  Dec(FLineCount);
  if FLineCount >= 1 then
    grdLines.RowCount := FLineCount + 1
  else
    grdLines.RowCount := 2;
  UpdateTotal;
end;

procedure TFrmBon.ExecLivraison(const ACode: string);
var
  R: Integer;
  Q: Integer;
  Code: string;
begin
  FQry.Close;
  FQry.SQL.Text :=
    'INSERT INTO BON_DE_LIVRAISON (Code_Bliv, Date_liv, Code_fourn) ' +
    'VALUES (:c, :d, :f)';
  FQry.Parameters.ParamByName('c').Value := ACode;
  FQry.Parameters.ParamByName('d').DataType := ftDateTime;
  FQry.Parameters.ParamByName('d').Value := DateOf(dtpDate.Date);
  FQry.Parameters.ParamByName('f').Value := FTiersCodes[cboTiers.ItemIndex];
  FQry.ExecSQL;

  for R := 1 to FLineCount do
  begin
    Code := grdLines.Cells[0, R];
    Q := StrToInt(grdLines.Cells[2, R]);

    FQry.Close;
    FQry.SQL.Text :=
      'INSERT INTO CONTIENT_LIVRAISON (Code_art, Code_Bliv, Qtt_liv, Prix) ' +
      'VALUES (:a, :b, :q, :p)';
    FQry.Parameters.ParamByName('a').Value := Code;
    FQry.Parameters.ParamByName('b').Value := ACode;
    FQry.Parameters.ParamByName('q').Value := Q;
    FQry.Parameters.ParamByName('p').DataType := ftCurrency;
    FQry.Parameters.ParamByName('p').Value := StrToCurr(grdLines.Cells[3, R]);
    FQry.ExecSQL;

    FQry.Close;
    FQry.SQL.Text :=
      'UPDATE ARTICLE SET Qtt_stock = IIf(Qtt_stock IS NULL, 0, Qtt_stock) + :q ' +
      'WHERE Code_art = :c';
    FQry.Parameters.ParamByName('q').Value := Q;
    FQry.Parameters.ParamByName('c').Value := Code;
    FQry.ExecSQL;
    if FQry.RowsAffected = 0 then
      raise Exception.Create('Article introuvable : ' + Code);
  end;
end;

procedure TFrmBon.ExecAffectation(const ACode: string);
var
  R: Integer;
  Q: Integer;
  Code: string;
begin
  FQry.Close;
  FQry.SQL.Text :=
    'INSERT INTO BON_D_AFFECTATION (Code_Aff, Date_Aff, Code_serv) ' +
    'VALUES (:c, :d, :s)';
  FQry.Parameters.ParamByName('c').Value := ACode;
  FQry.Parameters.ParamByName('d').DataType := ftDateTime;
  FQry.Parameters.ParamByName('d').Value := DateOf(dtpDate.Date);
  FQry.Parameters.ParamByName('s').Value := FTiersCodes[cboTiers.ItemIndex];
  FQry.ExecSQL;

  for R := 1 to FLineCount do
  begin
    Code := grdLines.Cells[0, R];
    Q := StrToInt(grdLines.Cells[2, R]);

    FQry.Close;
    FQry.SQL.Text :=
      'INSERT INTO CONTIENT_AFFECTATION (Code_art, Code_Aff, Qtt_aff) ' +
      'VALUES (:a, :b, :q)';
    FQry.Parameters.ParamByName('a').Value := Code;
    FQry.Parameters.ParamByName('b').Value := ACode;
    FQry.Parameters.ParamByName('q').Value := Q;
    FQry.ExecSQL;

    { Le contrôle du stock est fait par la requête elle-même, au moment de
      l'enregistrement : le stock lu à l'écran a pu changer entre-temps. }
    FQry.Close;
    FQry.SQL.Text :=
      'UPDATE ARTICLE SET Qtt_stock = Qtt_stock - :q ' +
      'WHERE Code_art = :c AND Qtt_stock >= :q2';
    FQry.Parameters.ParamByName('q').Value := Q;
    FQry.Parameters.ParamByName('c').Value := Code;
    FQry.Parameters.ParamByName('q2').Value := Q;
    FQry.ExecSQL;
    if FQry.RowsAffected = 0 then
      raise Exception.Create('Stock insuffisant (ou article introuvable) : ' + Code);
  end;
end;

{ Un bon de commande est une demande d'un service : il n'entre ni ne sort rien
  du stock, il est seulement enregistré. }
procedure TFrmBon.ExecCommande(const ACode: string);
var
  R: Integer;
begin
  FQry.Close;
  FQry.SQL.Text :=
    'INSERT INTO BON_DE_COMMANDE (Code_BC, Date_BC, Code_serv) ' +
    'VALUES (:c, :d, :s)';
  FQry.Parameters.ParamByName('c').Value := ACode;
  FQry.Parameters.ParamByName('d').DataType := ftDateTime;
  FQry.Parameters.ParamByName('d').Value := DateOf(dtpDate.Date);
  FQry.Parameters.ParamByName('s').Value := FTiersCodes[cboTiers.ItemIndex];
  FQry.ExecSQL;

  for R := 1 to FLineCount do
  begin
    FQry.Close;
    FQry.SQL.Text :=
      'INSERT INTO COMMANDER (Code_art, Code_BC, Qtt_cm) VALUES (:a, :b, :q)';
    FQry.Parameters.ParamByName('a').Value := grdLines.Cells[0, R];
    FQry.Parameters.ParamByName('b').Value := ACode;
    FQry.Parameters.ParamByName('q').Value := StrToInt(grdLines.Cells[2, R]);
    FQry.ExecSQL;
  end;
end;

procedure TFrmBon.SaveClick(Sender: TObject);
var
  Code: string;
begin
  if cboTiers.ItemIndex < 0 then
  begin
    if FKind = bkLivraison then
      raise Exception.Create('Choisissez le fournisseur.')
    else
      raise Exception.Create('Choisissez le service destinataire.');
  end;
  if FLineCount = 0 then
    raise Exception.Create('Le bon ne contient aucune ligne.');

  { Recalculé au dernier moment : un autre poste a pu prendre ce numéro. }
  Code := NextCode;
  edtCode.Text := Code;

  FConn.BeginTrans;
  try
    case FKind of
      bkLivraison: ExecLivraison(Code);
      bkAffectation: ExecAffectation(Code);
    else
      ExecCommande(Code);
    end;
    FConn.CommitTrans;
  except
    if FConn.InTransaction then
      FConn.RollbackTrans;
    LoadArticles;
    raise;
  end;

  if FKind = bkCommande then
    ShowMessage('Bon ' + Code + ' enregistré. Le stock n''est pas modifié.')
  else
    ShowMessage('Bon ' + Code + ' enregistré. Le stock a été mis à jour.');
  LoadArticles;
  NewBon;
end;

procedure TFrmBon.CloseClick(Sender: TObject);
begin
  Close;
end;

end.
