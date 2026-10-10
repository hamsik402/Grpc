unit UnitStock;

{ État du stock avec alertes de seuil.
  - Rupture        : quantité <= 0
  - Sous le seuil  : quantité <= seuil d'alerte de l'article
  Filtre par texte ou « alertes seulement », modification du seuil d'un
  article, impression de la liste affichée. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Types,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.Grids, Vcl.Dialogs, Vcl.Printers,
  Vcl.Graphics, Data.DB, Data.Win.ADODB, UnitMouv;

type
  TStockItem = record
    Code: string;
    Designation: string;
    Typ: string;
    Qtt: Integer;
    Seuil: Integer;
    HasSeuil: Boolean;
  end;

  TFrmStock = class(TForm)
  private
    FConn: TADOConnection;
    FQry: TADOQuery;
    FItems: array of TStockItem;
    FView: array of Integer;   { indices de FItems affichés (après filtre) }
    grdStock: TStringGrid;
    edtSearch: TEdit;
    chkAlerts: TCheckBox;
    edtSeuil: TEdit;
    lblSummary: TLabel;

    function StatusOf(const AItem: TStockItem): string;
    function AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
    function AddButton(const ACaption: string; ALeft, ATop, AWidth: Integer;
      AHandler: TNotifyEvent): TButton;
    procedure BuildUI;
    procedure LoadItems;
    procedure ApplyFilter(const AKeepCode: string);
    function SelectedCode: string;
    procedure PrintList;

    procedure FilterChange(Sender: TObject);
    procedure GridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure GridDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure SeuilClick(Sender: TObject);
    procedure MouvClick(Sender: TObject);
    procedure PrintClick(Sender: TObject);
    procedure CloseClick(Sender: TObject);
  public
    constructor CreateStock(AOwner: TComponent; AConn: TADOConnection);
  end;

procedure OpenStock(AConn: TADOConnection);

implementation

const
  ST_OK = 'OK';
  ST_LOW = 'Sous le seuil';
  ST_OUT = 'Rupture';

procedure OpenStock(AConn: TADOConnection);
var
  F: TFrmStock;
begin
  F := TFrmStock.CreateStock(Application, AConn);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

{ TFrmStock }

constructor TFrmStock.CreateStock(AOwner: TComponent; AConn: TADOConnection);
begin
  inherited CreateNew(AOwner);
  FConn := AConn;
  FQry := TADOQuery.Create(Self);
  FQry.Connection := FConn;
  BuildUI;
  LoadItems;
  ApplyFilter('');
end;

function TFrmStock.StatusOf(const AItem: TStockItem): string;
begin
  if AItem.Qtt <= 0 then
    Result := ST_OUT
  else if AItem.HasSeuil and (AItem.Qtt <= AItem.Seuil) then
    Result := ST_LOW
  else
    Result := ST_OK;
end;

function TFrmStock.AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
begin
  Result := TLabel.Create(Self);
  Result.Parent := Self;
  Result.Left := ALeft;
  Result.Top := ATop;
  Result.Caption := ACaption;
end;

function TFrmStock.AddButton(const ACaption: string; ALeft, ATop,
  AWidth: Integer; AHandler: TNotifyEvent): TButton;
begin
  Result := TButton.Create(Self);
  Result.Parent := Self;
  Result.SetBounds(ALeft, ATop, AWidth, 28);
  Result.Caption := ACaption;
  Result.OnClick := AHandler;
end;

procedure TFrmStock.BuildUI;
begin
  BorderStyle := bsDialog;
  Position := poOwnerFormCenter;
  ClientWidth := 800;
  ClientHeight := 580;
  Caption := 'État du stock';

  AddLabel('Rechercher (code ou désignation)', 16, 12);
  edtSearch := TEdit.Create(Self);
  edtSearch.Parent := Self;
  edtSearch.SetBounds(16, 30, 260, 21);
  edtSearch.OnChange := FilterChange;

  chkAlerts := TCheckBox.Create(Self);
  chkAlerts.Parent := Self;
  chkAlerts.SetBounds(300, 32, 200, 17);
  chkAlerts.Caption := 'Afficher les alertes seulement';
  chkAlerts.OnClick := FilterChange;

  grdStock := TStringGrid.Create(Self);
  grdStock.Parent := Self;
  grdStock.SetBounds(16, 62, 768, 410);
  grdStock.FixedCols := 0;
  grdStock.FixedRows := 1;
  grdStock.RowCount := 2;
  grdStock.ColCount := 6;
  grdStock.ColWidths[0] := 100;
  grdStock.ColWidths[1] := 300;
  grdStock.ColWidths[2] := 110;
  grdStock.ColWidths[3] := 70;
  grdStock.ColWidths[4] := 70;
  grdStock.ColWidths[5] := 100;
  grdStock.Options := grdStock.Options + [goRowSelect] - [goEditing];
  grdStock.Cells[0, 0] := 'Code';
  grdStock.Cells[1, 0] := 'Désignation';
  grdStock.Cells[2, 0] := 'Type';
  grdStock.Cells[3, 0] := 'Stock';
  grdStock.Cells[4, 0] := 'Seuil';
  grdStock.Cells[5, 0] := 'Statut';
  grdStock.OnSelectCell := GridSelectCell;
  grdStock.OnDrawCell := GridDrawCell;

  lblSummary := AddLabel('', 16, 482);

  AddLabel('Seuil d''alerte de l''article sélectionné', 16, 512);
  edtSeuil := TEdit.Create(Self);
  edtSeuil.Parent := Self;
  edtSeuil.SetBounds(250, 508, 70, 21);
  AddButton('Modifier le seuil', 332, 504, 130, SeuilClick);

  AddButton('Mouvements de l''article', 380, 540, 170, MouvClick);
  AddButton('Imprimer', 560, 540, 110, PrintClick);
  AddButton('Fermer', 684, 540, 100, CloseClick);
end;

procedure TFrmStock.LoadItems;
var
  N: Integer;
begin
  SetLength(FItems, 0);
  FQry.Close;
  FQry.SQL.Text :=
    'SELECT Code_art, Designation, Type_art, Qtt_stock, Seuil_alerte ' +
    'FROM ARTICLE ORDER BY Designation';
  FQry.Open;
  while not FQry.Eof do
  begin
    N := Length(FItems);
    SetLength(FItems, N + 1);
    FItems[N].Code := FQry.FieldByName('Code_art').AsString;
    FItems[N].Designation := FQry.FieldByName('Designation').AsString;
    FItems[N].Typ := FQry.FieldByName('Type_art').AsString;
    FItems[N].Qtt := FQry.FieldByName('Qtt_stock').AsInteger;
    FItems[N].HasSeuil := not FQry.FieldByName('Seuil_alerte').IsNull;
    FItems[N].Seuil := FQry.FieldByName('Seuil_alerte').AsInteger;
    FQry.Next;
  end;
  FQry.Close;
end;

procedure TFrmStock.ApplyFilter(const AKeepCode: string);
var
  I, R, NLow, NOut, SelRow: Integer;
  Needle, St: string;
begin
  Needle := LowerCase(Trim(edtSearch.Text));
  SetLength(FView, 0);
  NLow := 0;
  NOut := 0;
  for I := 0 to High(FItems) do
  begin
    St := StatusOf(FItems[I]);
    if St = ST_LOW then
      Inc(NLow)
    else if St = ST_OUT then
      Inc(NOut);
  end;

  for I := 0 to High(FItems) do
  begin
    if chkAlerts.Checked and (StatusOf(FItems[I]) = ST_OK) then
      Continue;
    if (Needle <> '') and
       (Pos(Needle, LowerCase(FItems[I].Code)) = 0) and
       (Pos(Needle, LowerCase(FItems[I].Designation)) = 0) then
      Continue;
    SetLength(FView, Length(FView) + 1);
    FView[High(FView)] := I;
  end;

  SelRow := 1;
  if Length(FView) = 0 then
  begin
    for R := 0 to grdStock.ColCount - 1 do
      grdStock.Cells[R, 1] := '';
    grdStock.RowCount := 2;
  end
  else
  begin
    grdStock.RowCount := Length(FView) + 1;
    for R := 0 to High(FView) do
    begin
      with FItems[FView[R]] do
      begin
        grdStock.Cells[0, R + 1] := Code;
        grdStock.Cells[1, R + 1] := Designation;
        grdStock.Cells[2, R + 1] := Typ;
        grdStock.Cells[3, R + 1] := IntToStr(Qtt);
        if HasSeuil then
          grdStock.Cells[4, R + 1] := IntToStr(Seuil)
        else
          grdStock.Cells[4, R + 1] := '';
        grdStock.Cells[5, R + 1] := StatusOf(FItems[FView[R]]);
        if (AKeepCode <> '') and SameText(Code, AKeepCode) then
          SelRow := R + 1;
      end;
    end;
  end;
  grdStock.Row := SelRow;
  grdStock.Invalidate;

  lblSummary.Caption := Format(
    '%d article(s) au total - %d sous le seuil - %d en rupture. Affichés : %d.',
    [Length(FItems), NLow, NOut, Length(FView)]);
  if (SelRow >= 1) and (SelRow <= Length(FView)) then
    edtSeuil.Text := grdStock.Cells[4, SelRow]
  else
    edtSeuil.Text := '';
end;

function TFrmStock.SelectedCode: string;
var
  R: Integer;
begin
  Result := '';
  R := grdStock.Row;
  if (R >= 1) and (R <= Length(FView)) then
    Result := grdStock.Cells[0, R];
end;

procedure TFrmStock.FilterChange(Sender: TObject);
begin
  ApplyFilter('');
end;

procedure TFrmStock.GridSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  if (ARow >= 1) and (ARow <= Length(FView)) then
    edtSeuil.Text := grdStock.Cells[4, ARow]
  else
    edtSeuil.Text := '';
end;

procedure TFrmStock.GridDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  St: string;
  Bg: TColor;
begin
  if (ARow < 1) or (ARow > Length(FView)) or (gdSelected in State) then
    Exit;
  St := grdStock.Cells[5, ARow];
  if St = ST_OUT then
    Bg := RGB(255, 190, 190)
  else if St = ST_LOW then
    Bg := RGB(255, 228, 160)
  else
    Exit;
  with grdStock.Canvas do
  begin
    Brush.Color := Bg;
    FillRect(Rect);
    Font.Color := clBlack;
    TextOut(Rect.Left + 4, Rect.Top + 2, grdStock.Cells[ACol, ARow]);
  end;
end;

procedure TFrmStock.SeuilClick(Sender: TObject);
var
  Code: string;
  V: Integer;
begin
  Code := SelectedCode;
  if Code = '' then
    raise Exception.Create('Sélectionnez un article dans la liste.');
  if not TryStrToInt(Trim(edtSeuil.Text), V) or (V < 0) then
    raise Exception.Create('Le seuil doit être un entier positif ou nul.');
  FQry.Close;
  FQry.SQL.Text := 'UPDATE ARTICLE SET Seuil_alerte = :s WHERE Code_art = :c';
  FQry.Parameters.ParamByName('s').Value := V;
  FQry.Parameters.ParamByName('c').Value := Code;
  FQry.ExecSQL;
  LoadItems;
  ApplyFilter(Code);
end;

procedure TFrmStock.MouvClick(Sender: TObject);
begin
  if SelectedCode = '' then
    raise Exception.Create('Sélectionnez un article dans la liste.');
  OpenMouvements(FConn, SelectedCode);
end;

procedure TFrmStock.CloseClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmStock.PrintClick(Sender: TObject);
begin
  if Length(FView) = 0 then
    raise Exception.Create('Aucune ligne à imprimer.');
  if Printer.Printers.Count = 0 then
    raise Exception.Create('Aucune imprimante n''est installée.');
  with TPrintDialog.Create(Self) do
  try
    if not Execute then
      Exit;
  finally
    Free;
  end;
  PrintList;
end;

procedure TFrmStock.PrintList;
var
  Cv: TCanvas;
  Margin, Y, LineH, PageW, TableW, I, K: Integer;
  ColX, ColW: array[0..4] of Integer;
  Heads: array[0..4] of string;
  Right: array[0..4] of Boolean;
  Sub: string;

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
    for J := 0 to 4 do
      Cell(J, Heads[J]);
    Inc(Y, LineH);
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    Cv.Font.Style := [];
    Inc(Y, 4);
  end;

begin
  Cv := Printer.Canvas;
  Printer.Title := 'Etat du stock';
  Printer.BeginDoc;
  try
    Cv.Font.Name := 'Arial';
    Cv.Font.Size := 10;
    Margin := Printer.PageWidth div 14;
    PageW := Printer.PageWidth;
    TableW := PageW - 2 * Margin;
    LineH := Cv.TextHeight('Ag') * 3 div 2;
    Y := Margin;

    Heads[0] := 'Code';  Heads[1] := 'Désignation'; Heads[2] := 'Stock';
    Heads[3] := 'Seuil'; Heads[4] := 'Statut';
    Right[0] := False; Right[1] := False; Right[2] := True;
    Right[3] := True;  Right[4] := False;
    ColW[0] := TableW * 14 div 100;
    ColW[1] := TableW * 42 div 100;
    ColW[2] := TableW * 11 div 100;
    ColW[3] := TableW * 11 div 100;
    ColW[4] := TableW - ColW[0] - ColW[1] - ColW[2] - ColW[3];
    ColX[0] := Margin;
    for K := 1 to 4 do
      ColX[K] := ColX[K - 1] + ColW[K - 1];

    Cv.Font.Style := [fsBold];
    Cv.Font.Size := 12;
    Cv.TextOut(Margin, Y, 'NAFTAL - District de Bouira');
    Inc(Y, LineH * 2);
    Cv.Font.Size := 14;
    Cv.TextOut(Margin, Y, 'ÉTAT DU STOCK');
    Inc(Y, LineH * 2);
    Cv.Font.Size := 10;
    Cv.Font.Style := [];
    Sub := 'Situation au ' + FormatDateTime('dd/mm/yyyy', Now);
    if chkAlerts.Checked then
      Sub := Sub + ' - articles en alerte seulement';
    if Trim(edtSearch.Text) <> '' then
      Sub := Sub + ' - filtre : « ' + Trim(edtSearch.Text) + ' »';
    Cv.TextOut(Margin, Y, Sub);
    Inc(Y, LineH * 2);

    TableHeader;
    for I := 1 to Length(FView) do
    begin
      if Y + LineH * 2 > Printer.PageHeight - Margin then
      begin
        Printer.NewPage;
        Y := Margin;
        TableHeader;
      end;
      Cell(0, grdStock.Cells[0, I]);
      Cell(1, grdStock.Cells[1, I]);
      Cell(2, grdStock.Cells[3, I]);
      Cell(3, grdStock.Cells[4, I]);
      if grdStock.Cells[5, I] <> ST_OK then
        Cv.Font.Style := [fsBold];
      Cell(4, grdStock.Cells[5, I]);
      Cv.Font.Style := [];
      Inc(Y, LineH);
    end;
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    Inc(Y, LineH div 2);
    Cv.Font.Style := [fsBold];
    Cv.TextOut(Margin, Y, lblSummary.Caption);

    Printer.EndDoc;
  except
    Printer.Abort;
    raise;
  end;
end;

end.
