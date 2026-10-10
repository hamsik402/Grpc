unit UnitBon;

{ Bons de livraison (entrée de stock, fournisseur -> magasin) et
  bons d'affectation (sortie de stock, magasin -> service).
  Une seule fiche sert aux deux types de bon ; elle est construite en code.
  L'enregistrement d'un bon (en-tête + lignes + mise à jour du stock) se fait
  dans une seule transaction : tout est enregistré, ou rien.

  Correction d'un bon (AEditCode renseigné) : le bon d'origine est marqué
  « Annulé » (motif : remplacé par le nouveau numéro) et un nouveau bon est
  créé avec les données corrigées, dans la même transaction. Le stock reçoit
  la différence entre l'ancien et le nouveau bon. Une commande corrigée
  reprend les livraisons qui lui étaient rattachées. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.DateUtils, System.Variants,
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
    FCmdCodes: TStringList;   { commandes proposées ; '' = aucune }
    FArtCodes: TStringList;
    FArtStock: array of Integer;
    FLineCount: Integer;
    FEditCode: string;        { bon corrigé ; '' = nouveau bon }
    FOldCodes: TStringList;   { lignes du bon corrigé (articles) }
    FOldQtys: array of Integer;

    edtCode: TEdit;
    dtpDate: TDateTimePicker;
    cboTiers: TComboBox;
    cboArticle: TComboBox;
    cboCmd: TComboBox;        { livraison seulement : commande liée }
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
    procedure LoadCommandes;
    procedure LoadArticles;
    function NextCode: string;
    procedure NewBon;
    procedure BonTable(out ATbl, AFld: string);
    procedure LoadForEdit(const ACode: string);
    function OldQtyOf(const ACode: string): Integer;
    procedure MarkOldReplaced(const ANewCode: string);
    procedure ReverseOld;
    procedure RepointLivraisons(const ANewCode: string);

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
      AKind: TBonKind; const AEditCode: string = '');
    destructor Destroy; override;
  end;

procedure OpenBon(AConn: TADOConnection; AKind: TBonKind;
  const AEditCode: string = '');

implementation

procedure OpenBon(AConn: TADOConnection; AKind: TBonKind;
  const AEditCode: string);
var
  F: TFrmBon;
begin
  F := TFrmBon.CreateBon(Application, AConn, AKind, AEditCode);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

{ TFrmBon }

constructor TFrmBon.CreateBon(AOwner: TComponent; AConn: TADOConnection;
  AKind: TBonKind; const AEditCode: string);
begin
  inherited CreateNew(AOwner);
  FConn := AConn;
  FKind := AKind;
  FTiersCodes := TStringList.Create;
  FCmdCodes := TStringList.Create;
  FOldCodes := TStringList.Create;
  FArtCodes := TStringList.Create;
  FQry := TADOQuery.Create(Self);
  FQry.Connection := FConn;
  BuildUI;
  OnShow := FormShowHandler;
  LoadTiers;
  LoadArticles;
  if FKind = bkLivraison then
    LoadCommandes;
  NewBon;
  if AEditCode <> '' then
    LoadForEdit(AEditCode);
end;

destructor TFrmBon.Destroy;
begin
  FTiersCodes.Free;
  FCmdCodes.Free;
  FOldCodes.Free;
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

  if IsLiv then
  begin
    AddLabel('Commande liée (facultatif) :', 300, 126);
    cboCmd := TComboBox.Create(Self);
    cboCmd.Parent := Self;
    cboCmd.Style := csDropDownList;
    cboCmd.SetBounds(450, 121, 284, 21);
  end;

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

{ Commandes non annulées proposées au rattachement d'une livraison. }
procedure TFrmBon.LoadCommandes;
var
  Txt: string;
begin
  cboCmd.Items.Clear;
  FCmdCodes.Clear;
  cboCmd.Items.Add('(aucune)');
  FCmdCodes.Add('');
  FQry.Close;
  FQry.SQL.Text :=
    'SELECT b.Code_BC, b.Date_BC, s.Nom_serv FROM BON_DE_COMMANDE AS b ' +
    'LEFT JOIN SERVICE_NAFTAL AS s ON b.Code_serv = s.Code_serv ' +
    'WHERE b.Annule IS NULL OR b.Annule = False ' +
    'ORDER BY b.Date_BC DESC, b.Code_BC DESC';
  FQry.Open;
  while not FQry.Eof do
  begin
    Txt := FQry.Fields[0].AsString;
    if not FQry.Fields[1].IsNull then
      Txt := Txt + ' du ' + FormatDateTime('dd/mm/yyyy', FQry.Fields[1].AsDateTime);
    Txt := Txt + ' - ' + FQry.Fields[2].AsString;
    cboCmd.Items.Add(Txt);
    FCmdCodes.Add(FQry.Fields[0].AsString);
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

procedure TFrmBon.BonTable(out ATbl, AFld: string);
begin
  case FKind of
    bkLivraison:
      begin
        ATbl := 'BON_DE_LIVRAISON';
        AFld := 'Code_Bliv';
      end;
    bkAffectation:
      begin
        ATbl := 'BON_D_AFFECTATION';
        AFld := 'Code_Aff';
      end;
  else
    begin
      ATbl := 'BON_DE_COMMANDE';
      AFld := 'Code_BC';
    end;
  end;
end;

function TFrmBon.OldQtyOf(const ACode: string): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to FOldCodes.Count - 1 do
    if SameText(FOldCodes[I], ACode) then
      Inc(Result, FOldQtys[I]);
end;

{ Charge dans la fiche le bon à corriger (en-tête et lignes). }
procedure TFrmBon.LoadForEdit(const ACode: string);
var
  Tbl, Fld, LinesTbl, QtyFld, TiersFld, DateFld, Extra, SQLLines: string;
  Idx, R, N: Integer;
  Q: Integer;
begin
  BonTable(Tbl, Fld);
  case FKind of
    bkLivraison:
      begin
        DateFld := 'Date_liv';
        TiersFld := 'Code_fourn';
        LinesTbl := 'CONTIENT_LIVRAISON';
        QtyFld := 'Qtt_liv';
      end;
    bkAffectation:
      begin
        DateFld := 'Date_Aff';
        TiersFld := 'Code_serv';
        LinesTbl := 'CONTIENT_AFFECTATION';
        QtyFld := 'Qtt_aff';
      end;
  else
    begin
      DateFld := 'Date_BC';
      TiersFld := 'Code_serv';
      LinesTbl := 'COMMANDER';
      QtyFld := 'Qtt_cm';
    end;
  end;
  Extra := '';
  if FKind = bkLivraison then
    Extra := ', Code_BC AS Cmd';

  FQry.Close;
  FQry.SQL.Text := 'SELECT ' + DateFld + ' AS Dt, ' + TiersFld + ' AS Tiers, ' +
    'Annule' + Extra + ' FROM ' + Tbl + ' WHERE ' + Fld + ' = :c';
  FQry.Parameters.ParamByName('c').Value := ACode;
  FQry.Open;
  try
    if FQry.Eof then
      raise Exception.Create('Bon introuvable : ' + ACode);
    if FQry.FieldByName('Annule').AsBoolean then
      raise Exception.Create('Un bon annulé ne peut pas être corrigé.');
    if not FQry.FieldByName('Dt').IsNull then
      dtpDate.Date := FQry.FieldByName('Dt').AsDateTime;
    cboTiers.ItemIndex := FTiersCodes.IndexOf(FQry.FieldByName('Tiers').AsString);
    if FKind = bkLivraison then
    begin
      Idx := FCmdCodes.IndexOf(FQry.FieldByName('Cmd').AsString);
      if Idx < 0 then
        Idx := 0;
      cboCmd.ItemIndex := Idx;
    end;
  finally
    FQry.Close;
  end;

  SQLLines := 'SELECT c.Code_art, a.Designation, c.' + QtyFld + ' AS Qtt';
  if FKind = bkLivraison then
    SQLLines := SQLLines + ', c.Prix';
  SQLLines := SQLLines + ' FROM ' + LinesTbl + ' AS c LEFT JOIN ARTICLE AS a ' +
    'ON c.Code_art = a.Code_art WHERE c.' + Fld + ' = :c ORDER BY c.Code_art';
  FQry.SQL.Text := SQLLines;
  FQry.Parameters.ParamByName('c').Value := ACode;
  FQry.Open;
  FOldCodes.Clear;
  SetLength(FOldQtys, 0);
  R := 0;
  while not FQry.Eof do
  begin
    Inc(R);
    if grdLines.RowCount < R + 1 then
      grdLines.RowCount := R + 1;
    Q := FQry.FieldByName('Qtt').AsInteger;
    grdLines.Cells[0, R] := FQry.FieldByName('Code_art').AsString;
    grdLines.Cells[1, R] := FQry.FieldByName('Designation').AsString;
    grdLines.Cells[2, R] := IntToStr(Q);
    if FKind = bkLivraison then
      grdLines.Cells[3, R] := CurrToStr(FQry.FieldByName('Prix').AsCurrency);
    FOldCodes.Add(FQry.FieldByName('Code_art').AsString);
    N := Length(FOldQtys);
    SetLength(FOldQtys, N + 1);
    FOldQtys[N] := Q;
    FQry.Next;
  end;
  FQry.Close;
  FLineCount := R;
  if FLineCount > 0 then
    grdLines.RowCount := FLineCount + 1;

  FEditCode := ACode;
  Caption := Caption + ' - correction du bon ' + ACode;
  btnSave.Caption := 'Enregistrer la correction';
  UpdateTotal;
end;

{ Marque le bon d'origine comme annulé, remplacé par le nouveau. }
procedure TFrmBon.MarkOldReplaced(const ANewCode: string);
var
  Tbl, Fld: string;
begin
  BonTable(Tbl, Fld);
  FQry.Close;
  FQry.SQL.Text := 'UPDATE ' + Tbl +
    ' SET Annule = :a, Date_annul = :d, Motif_annul = :m' +
    ' WHERE ' + Fld + ' = :c AND (Annule IS NULL OR Annule = False)';
  FQry.Parameters.ParamByName('a').DataType := ftBoolean;
  FQry.Parameters.ParamByName('a').Value := True;
  FQry.Parameters.ParamByName('d').DataType := ftDateTime;
  FQry.Parameters.ParamByName('d').Value := Now;
  FQry.Parameters.ParamByName('m').Value := 'Corrigé : remplacé par ' + ANewCode;
  FQry.Parameters.ParamByName('c').Value := FEditCode;
  FQry.ExecSQL;
  if FQry.RowsAffected = 0 then
    raise Exception.Create('Le bon à corriger a été annulé ou modifié entre-temps.');
end;

{ Retire du stock l'effet du bon remplacé (livraison : on retire ce qui était
  entré ; affectation : on rend ce qui était sorti). }
procedure TFrmBon.ReverseOld;
var
  I: Integer;
begin
  if FKind = bkCommande then
    Exit;
  for I := 0 to FOldCodes.Count - 1 do
  begin
    FQry.Close;
    if FKind = bkLivraison then
    begin
      FQry.SQL.Text :=
        'UPDATE ARTICLE SET Qtt_stock = Qtt_stock - :q ' +
        'WHERE Code_art = :c AND Qtt_stock >= :q2';
      FQry.Parameters.ParamByName('q').Value := FOldQtys[I];
      FQry.Parameters.ParamByName('c').Value := FOldCodes[I];
      FQry.Parameters.ParamByName('q2').Value := FOldQtys[I];
      FQry.ExecSQL;
      if FQry.RowsAffected = 0 then
        raise Exception.CreateFmt(
          'Correction impossible : le stock de l''article %s deviendrait négatif ' +
          '(des articles de ce bon ont déjà été sortis).', [FOldCodes[I]]);
    end
    else
    begin
      FQry.SQL.Text :=
        'UPDATE ARTICLE SET Qtt_stock = ' +
        'IIf(Qtt_stock IS NULL, 0, Qtt_stock) + :q WHERE Code_art = :c';
      FQry.Parameters.ParamByName('q').Value := FOldQtys[I];
      FQry.Parameters.ParamByName('c').Value := FOldCodes[I];
      FQry.ExecSQL;
      if FQry.RowsAffected = 0 then
        raise Exception.Create('Article introuvable : ' + FOldCodes[I]);
    end;
  end;
end;

{ Commande corrigée : les livraisons déjà rattachées suivent le nouveau bon. }
procedure TFrmBon.RepointLivraisons(const ANewCode: string);
begin
  FQry.Close;
  FQry.SQL.Text := 'UPDATE BON_DE_LIVRAISON SET Code_BC = :n WHERE Code_BC = :o';
  FQry.Parameters.ParamByName('n').Value := ANewCode;
  FQry.Parameters.ParamByName('o').Value := FEditCode;
  FQry.ExecSQL;
end;

procedure TFrmBon.NewBon;
var
  R: Integer;
begin
  edtCode.Text := NextCode;
  dtpDate.Date := Date;
  cboTiers.ItemIndex := -1;
  cboArticle.ItemIndex := -1;
  if Assigned(cboCmd) and (cboCmd.Items.Count > 0) then
    cboCmd.ItemIndex := 0;
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
  { en correction, les quantités du bon remplacé sont rendues au stock }
  if (FKind = bkAffectation) and (Q > FArtStock[Idx] + OldQtyOf(Code)) then
    raise Exception.CreateFmt('Stock insuffisant. Disponible : %d',
      [FArtStock[Idx] + OldQtyOf(Code)]);

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
    'INSERT INTO BON_DE_LIVRAISON (Code_Bliv, Date_liv, Code_fourn, Code_BC) ' +
    'VALUES (:c, :d, :f, :bc)';
  FQry.Parameters.ParamByName('c').Value := ACode;
  FQry.Parameters.ParamByName('d').DataType := ftDateTime;
  FQry.Parameters.ParamByName('d').Value := DateOf(dtpDate.Date);
  FQry.Parameters.ParamByName('f').Value := FTiersCodes[cboTiers.ItemIndex];
  FQry.Parameters.ParamByName('bc').DataType := ftString;
  if (cboCmd.ItemIndex <= 0) or (FCmdCodes[cboCmd.ItemIndex] = '') then
    FQry.Parameters.ParamByName('bc').Value := Null
  else
    FQry.Parameters.ParamByName('bc').Value := FCmdCodes[cboCmd.ItemIndex];
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
    if FEditCode <> '' then
    begin
      MarkOldReplaced(Code);
      { affectation : on rend d'abord l'ancien bon au stock, puis on retire le
        nouveau ; livraison : l'inverse, pour que la vérification du stock
        porte sur le solde final. }
      if FKind = bkAffectation then
        ReverseOld;
    end;
    case FKind of
      bkLivraison: ExecLivraison(Code);
      bkAffectation: ExecAffectation(Code);
    else
      ExecCommande(Code);
    end;
    if FEditCode <> '' then
    begin
      if FKind = bkLivraison then
        ReverseOld
      else if FKind = bkCommande then
        RepointLivraisons(Code);
    end;
    FConn.CommitTrans;
  except
    if FConn.InTransaction then
      FConn.RollbackTrans;
    LoadArticles;
    raise;
  end;

  if FEditCode <> '' then
  begin
    ShowMessage('Bon ' + FEditCode + ' corrigé : il est remplacé par le bon ' +
      Code + '. L''ancien bon est conservé, marqué « Annulé ».');
    ModalResult := mrOk;
    Exit;
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
