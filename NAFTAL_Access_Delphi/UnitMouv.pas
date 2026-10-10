unit UnitMouv;

{ Mouvements de stock d'un article : entrées (bons de livraison) et sorties
  (bons d'affectation), avec solde cumulé.

  Les mouvements sont reconstitués à partir des bons ; il n'y a pas de table
  de mouvements. Les modifications faites directement par « Entrée stock » /
  « Sortie stock » de l'écran principal, ou le stock saisi au départ, ne
  laissent pas de bon : elles apparaissent sur une première ligne
  « Stock initial / ajustements hors bons » = stock actuel - (entrées - sorties).
  Les bons annulés sont listés (en gris) mais ne comptent pas dans le solde. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.Grids, Vcl.Dialogs, Vcl.Printers,
  Vcl.Graphics, Data.DB, Data.Win.ADODB;

type
  TMouv = record
    Dt: TDateTime;
    DateTxt: string;
    Bon: string;
    IsEntree: Boolean;
    Tiers: string;
    Qtt: Integer;
    Annule: Boolean;
  end;

  TFrmMouv = class(TForm)
  private
    FConn: TADOConnection;
    FQry: TADOQuery;
    FArtCodes: TStringList;
    FMouv: array of TMouv;
    FStockActuel, FTotIn, FTotOut, FInitial: Integer;
    cboArticle: TComboBox;
    grdMouv: TStringGrid;
    lblSummary: TLabel;

    function AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
    function AddButton(const ACaption: string; ALeft, ATop, AWidth: Integer;
      AHandler: TNotifyEvent): TButton;
    procedure BuildUI;
    procedure LoadArticles;
    procedure CollectMouv(const ACode: string; AEntree: Boolean);
    procedure SortMouv;
    procedure LoadMouv;
    procedure PrintMouv;

    procedure ArticleChange(Sender: TObject);
    procedure GridDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure PrintClick(Sender: TObject);
    procedure CloseClick(Sender: TObject);
  public
    constructor CreateMouv(AOwner: TComponent; AConn: TADOConnection;
      const ACode: string);
    destructor Destroy; override;
  end;

procedure OpenMouvements(AConn: TADOConnection; const ACode: string = '');

implementation

procedure OpenMouvements(AConn: TADOConnection; const ACode: string);
var
  F: TFrmMouv;
begin
  F := TFrmMouv.CreateMouv(Application, AConn, ACode);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

{ TFrmMouv }

constructor TFrmMouv.CreateMouv(AOwner: TComponent; AConn: TADOConnection;
  const ACode: string);
var
  I: Integer;
begin
  inherited CreateNew(AOwner);
  FConn := AConn;
  FQry := TADOQuery.Create(Self);
  FQry.Connection := FConn;
  FArtCodes := TStringList.Create;
  BuildUI;
  LoadArticles;
  for I := 0 to FArtCodes.Count - 1 do
    if SameText(FArtCodes[I], ACode) then
      cboArticle.ItemIndex := I;
  LoadMouv;
end;

destructor TFrmMouv.Destroy;
begin
  FArtCodes.Free;
  inherited;
end;

function TFrmMouv.AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
begin
  Result := TLabel.Create(Self);
  Result.Parent := Self;
  Result.Left := ALeft;
  Result.Top := ATop;
  Result.Caption := ACaption;
end;

function TFrmMouv.AddButton(const ACaption: string; ALeft, ATop,
  AWidth: Integer; AHandler: TNotifyEvent): TButton;
begin
  Result := TButton.Create(Self);
  Result.Parent := Self;
  Result.SetBounds(ALeft, ATop, AWidth, 28);
  Result.Caption := ACaption;
  Result.OnClick := AHandler;
end;

procedure TFrmMouv.BuildUI;
begin
  BorderStyle := bsDialog;
  Position := poOwnerFormCenter;
  ClientWidth := 850;
  ClientHeight := 580;
  Caption := 'Mouvements de stock par article';

  AddLabel('Article', 16, 12);
  cboArticle := TComboBox.Create(Self);
  cboArticle.Parent := Self;
  cboArticle.Style := csDropDownList;
  cboArticle.SetBounds(16, 30, 500, 21);
  cboArticle.OnChange := ArticleChange;

  grdMouv := TStringGrid.Create(Self);
  grdMouv.Parent := Self;
  grdMouv.SetBounds(16, 62, 818, 400);
  grdMouv.FixedCols := 0;
  grdMouv.FixedRows := 1;
  grdMouv.RowCount := 2;
  grdMouv.ColCount := 7;
  grdMouv.ColWidths[0] := 80;
  grdMouv.ColWidths[1] := 110;
  grdMouv.ColWidths[2] := 190;
  grdMouv.ColWidths[3] := 190;
  grdMouv.ColWidths[4] := 70;
  grdMouv.ColWidths[5] := 70;
  grdMouv.ColWidths[6] := 70;
  grdMouv.Options := grdMouv.Options + [goRowSelect] - [goEditing];
  grdMouv.Cells[0, 0] := 'Date';
  grdMouv.Cells[1, 0] := 'N° du bon';
  grdMouv.Cells[2, 0] := 'Opération';
  grdMouv.Cells[3, 0] := 'Fournisseur / Service';
  grdMouv.Cells[4, 0] := 'Entrée';
  grdMouv.Cells[5, 0] := 'Sortie';
  grdMouv.Cells[6, 0] := 'Solde';
  grdMouv.OnDrawCell := GridDrawCell;

  lblSummary := AddLabel('', 16, 474);

  AddButton('Imprimer', 624, 540, 110, PrintClick);
  AddButton('Fermer', 744, 540, 90, CloseClick);
end;

procedure TFrmMouv.LoadArticles;
begin
  cboArticle.Items.Clear;
  FArtCodes.Clear;
  FQry.Close;
  FQry.SQL.Text := 'SELECT Code_art, Designation FROM ARTICLE ORDER BY Designation';
  FQry.Open;
  while not FQry.Eof do
  begin
    cboArticle.Items.Add(FQry.FieldByName('Code_art').AsString + ' - ' +
      FQry.FieldByName('Designation').AsString);
    FArtCodes.Add(FQry.FieldByName('Code_art').AsString);
    FQry.Next;
  end;
  FQry.Close;
end;

procedure TFrmMouv.CollectMouv(const ACode: string; AEntree: Boolean);
var
  N: Integer;
begin
  FQry.Close;
  if AEntree then
    FQry.SQL.Text :=
      'SELECT b.Date_liv AS Dt, b.Code_Bliv AS Bon, f.Nom AS Tiers, ' +
      'c.Qtt_liv AS Qtt, b.Annule AS Ann ' +
      'FROM (CONTIENT_LIVRAISON AS c INNER JOIN BON_DE_LIVRAISON AS b ' +
      'ON c.Code_Bliv = b.Code_Bliv) LEFT JOIN FOURNISSEUR AS f ' +
      'ON b.Code_fourn = f.Code_fourn WHERE c.Code_art = :a'
  else
    FQry.SQL.Text :=
      'SELECT b.Date_Aff AS Dt, b.Code_Aff AS Bon, s.Nom_serv AS Tiers, ' +
      'c.Qtt_aff AS Qtt, b.Annule AS Ann ' +
      'FROM (CONTIENT_AFFECTATION AS c INNER JOIN BON_D_AFFECTATION AS b ' +
      'ON c.Code_Aff = b.Code_Aff) LEFT JOIN SERVICE_NAFTAL AS s ' +
      'ON b.Code_serv = s.Code_serv WHERE c.Code_art = :a';
  FQry.Parameters.ParamByName('a').Value := ACode;
  FQry.Open;
  while not FQry.Eof do
  begin
    N := Length(FMouv);
    SetLength(FMouv, N + 1);
    FMouv[N].IsEntree := AEntree;
    FMouv[N].Bon := FQry.FieldByName('Bon').AsString;
    FMouv[N].Tiers := FQry.FieldByName('Tiers').AsString;
    FMouv[N].Qtt := FQry.FieldByName('Qtt').AsInteger;
    FMouv[N].Annule := FQry.FieldByName('Ann').AsBoolean;
    if FQry.FieldByName('Dt').IsNull then
    begin
      FMouv[N].Dt := 0;
      FMouv[N].DateTxt := '';
    end
    else
    begin
      FMouv[N].Dt := Trunc(FQry.FieldByName('Dt').AsDateTime);
      FMouv[N].DateTxt := FormatDateTime('dd/mm/yyyy', FMouv[N].Dt);
    end;
    FQry.Next;
  end;
  FQry.Close;
end;

{ Ordre chronologique ; à date égale, les entrées avant les sorties. }
procedure TFrmMouv.SortMouv;
var
  I, J: Integer;
  Tmp: TMouv;

  function Before(const A, B: TMouv): Boolean;
  begin
    if A.Dt <> B.Dt then
      Result := A.Dt < B.Dt
    else if A.IsEntree <> B.IsEntree then
      Result := A.IsEntree
    else
      Result := A.Bon < B.Bon;
  end;

begin
  for I := 1 to High(FMouv) do
  begin
    Tmp := FMouv[I];
    J := I - 1;
    while (J >= 0) and Before(Tmp, FMouv[J]) do
    begin
      FMouv[J + 1] := FMouv[J];
      Dec(J);
    end;
    FMouv[J + 1] := Tmp;
  end;
end;

procedure TFrmMouv.LoadMouv;
var
  Code, Op: string;
  I, K, Row, Bal: Integer;
begin
  SetLength(FMouv, 0);
  FStockActuel := 0;
  FTotIn := 0;
  FTotOut := 0;
  FInitial := 0;
  for I := 1 to grdMouv.RowCount - 1 do
    for K := 0 to grdMouv.ColCount - 1 do
      grdMouv.Cells[K, I] := '';
  grdMouv.RowCount := 2;

  if cboArticle.ItemIndex < 0 then
  begin
    lblSummary.Caption := 'Choisissez un article.';
    Exit;
  end;
  Code := FArtCodes[cboArticle.ItemIndex];

  FQry.Close;
  FQry.SQL.Text := 'SELECT Qtt_stock FROM ARTICLE WHERE Code_art = :a';
  FQry.Parameters.ParamByName('a').Value := Code;
  FQry.Open;
  FStockActuel := FQry.FieldByName('Qtt_stock').AsInteger;
  FQry.Close;

  CollectMouv(Code, True);
  CollectMouv(Code, False);
  SortMouv;

  for I := 0 to High(FMouv) do
    if not FMouv[I].Annule then
      if FMouv[I].IsEntree then
        Inc(FTotIn, FMouv[I].Qtt)
      else
        Inc(FTotOut, FMouv[I].Qtt);
  FInitial := FStockActuel - (FTotIn - FTotOut);

  Row := 0;
  Bal := FInitial;
  if FInitial <> 0 then
  begin
    Inc(Row);
    grdMouv.RowCount := Row + 1;
    grdMouv.Cells[2, Row] := 'Stock initial / ajustements hors bons';
    grdMouv.Cells[6, Row] := IntToStr(Bal);
  end;
  for I := 0 to High(FMouv) do
  begin
    Inc(Row);
    if grdMouv.RowCount < Row + 1 then
      grdMouv.RowCount := Row + 1;
    if FMouv[I].IsEntree then
      Op := 'Livraison'
    else
      Op := 'Affectation';
    if FMouv[I].Annule then
      Op := Op + ' (annulée)';
    grdMouv.Cells[0, Row] := FMouv[I].DateTxt;
    grdMouv.Cells[1, Row] := FMouv[I].Bon;
    grdMouv.Cells[2, Row] := Op;
    grdMouv.Cells[3, Row] := FMouv[I].Tiers;
    if FMouv[I].IsEntree then
      grdMouv.Cells[4, Row] := IntToStr(FMouv[I].Qtt)
    else
      grdMouv.Cells[5, Row] := IntToStr(FMouv[I].Qtt);
    if not FMouv[I].Annule then
    begin
      if FMouv[I].IsEntree then
        Inc(Bal, FMouv[I].Qtt)
      else
        Dec(Bal, FMouv[I].Qtt);
      grdMouv.Cells[6, Row] := IntToStr(Bal);
    end;
  end;

  lblSummary.Caption := Format(
    'Stock actuel : %d   Entrées : %d   Sorties : %d   Stock initial / ajustements hors bons : %d',
    [FStockActuel, FTotIn, FTotOut, FInitial]);
end;

procedure TFrmMouv.ArticleChange(Sender: TObject);
begin
  LoadMouv;
end;

procedure TFrmMouv.GridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if (ARow < 1) or (gdSelected in State) or
     (Pos('annulée', grdMouv.Cells[2, ARow]) = 0) then
    Exit;
  with grdMouv.Canvas do
  begin
    Brush.Color := RGB(220, 220, 220);
    FillRect(Rect);
    Font.Color := clGray;
    TextOut(Rect.Left + 4, Rect.Top + 2, grdMouv.Cells[ACol, ARow]);
  end;
end;

procedure TFrmMouv.CloseClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmMouv.PrintClick(Sender: TObject);
begin
  if cboArticle.ItemIndex < 0 then
    raise Exception.Create('Choisissez un article.');
  if Printer.Printers.Count = 0 then
    raise Exception.Create('Aucune imprimante n''est installée.');
  with TPrintDialog.Create(Self) do
  try
    if not Execute then
      Exit;
  finally
    Free;
  end;
  PrintMouv;
end;

procedure TFrmMouv.PrintMouv;
const
  NCOL = 7;
  Pct: array[0..NCOL - 1] of Integer = (11, 15, 22, 24, 9, 9, 10);
var
  Cv: TCanvas;
  Margin, Y, LineH, PageW, TableW, I, K, Used: Integer;
  ColX, ColW: array[0..NCOL - 1] of Integer;
  Heads: array[0..NCOL - 1] of string;
  Right: array[0..NCOL - 1] of Boolean;

  function Clip(const S: string; AWidth: Integer): string;
  begin
    Result := S;
    while (Result <> '') and (Cv.TextWidth(Result) > AWidth) do
      Delete(Result, Length(Result), 1);
  end;

  procedure Cell(ACol: Integer; const S: string);
  var
    T: string;
  begin
    T := Clip(S, ColW[ACol] - 8);
    if Right[ACol] then
      Cv.TextOut(ColX[ACol] + ColW[ACol] - 4 - Cv.TextWidth(T), Y, T)
    else
      Cv.TextOut(ColX[ACol] + 4, Y, T);
  end;

  procedure TableHeader;
  var
    J: Integer;
  begin
    Cv.Font.Style := [fsBold];
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    for J := 0 to NCOL - 1 do
      Cell(J, Heads[J]);
    Inc(Y, LineH);
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    Cv.Font.Style := [];
    Inc(Y, 4);
  end;

begin
  Cv := Printer.Canvas;
  Printer.Title := 'Mouvements article';
  Printer.BeginDoc;
  try
    Cv.Font.Name := 'Arial';
    Cv.Font.Size := 9;
    Margin := Printer.PageWidth div 16;
    PageW := Printer.PageWidth;
    TableW := PageW - 2 * Margin;
    LineH := Cv.TextHeight('Ag') * 3 div 2;
    Y := Margin;

    Heads[0] := 'Date';      Heads[1] := 'N° du bon'; Heads[2] := 'Opération';
    Heads[3] := 'Fourn. / Service'; Heads[4] := 'Entrée'; Heads[5] := 'Sortie';
    Heads[6] := 'Solde';
    for K := 0 to NCOL - 1 do
      Right[K] := K >= 4;
    Used := 0;
    for K := 0 to NCOL - 1 do
    begin
      if K = NCOL - 1 then
        ColW[K] := TableW - Used
      else
        ColW[K] := TableW * Pct[K] div 100;
      Inc(Used, ColW[K]);
    end;
    ColX[0] := Margin;
    for K := 1 to NCOL - 1 do
      ColX[K] := ColX[K - 1] + ColW[K - 1];

    Cv.Font.Style := [fsBold];
    Cv.Font.Size := 12;
    Cv.TextOut(Margin, Y, 'NAFTAL - District de Bouira');
    Inc(Y, LineH * 2);
    Cv.Font.Size := 14;
    Cv.TextOut(Margin, Y, 'MOUVEMENTS DE STOCK PAR ARTICLE');
    Inc(Y, LineH * 2);
    Cv.Font.Size := 10;
    Cv.TextOut(Margin, Y, Clip('Article : ' + cboArticle.Text, TableW));
    Inc(Y, LineH);
    Cv.Font.Style := [];
    Cv.TextOut(Margin, Y, 'Situation au ' + FormatDateTime('dd/mm/yyyy', Now));
    Inc(Y, LineH * 2);
    Cv.Font.Size := 9;

    TableHeader;
    for I := 1 to grdMouv.RowCount - 1 do
    begin
      if grdMouv.Cells[2, I] = '' then
        Continue;
      if Y + LineH * 2 > Printer.PageHeight - Margin then
      begin
        Printer.NewPage;
        Y := Margin;
        TableHeader;
      end;
      for K := 0 to NCOL - 1 do
        Cell(K, grdMouv.Cells[K, I]);
      Inc(Y, LineH);
    end;
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    Inc(Y, LineH div 2);
    Cv.Font.Style := [fsBold];
    Cv.TextOut(Margin, Y, Clip(lblSummary.Caption, TableW));

    Printer.EndDoc;
  except
    Printer.Abort;
    raise;
  end;
end;

end.
