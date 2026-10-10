unit UnitRappro;

{ Rapprochement commande / livraison.
  Une livraison peut être rattachée à un bon de commande (champ « Commande
  liée » du bon de livraison). Pour chaque commande non annulée, on compare
  par article la quantité commandée à la quantité reçue, en additionnant les
  livraisons liées non annulées (une commande peut être livrée en plusieurs
  fois). Aucun effet sur le stock : c'est un écran de suivi. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.Grids, Vcl.Dialogs, Vcl.Printers,
  Vcl.Graphics, Data.DB, Data.Win.ADODB;

type
  TQtyRow = record
    Bc: string;
    Art: string;
    Qty: Integer;
  end;

  TCmdRow = record
    Code: string;
    DateTxt: string;
    Service: string;
    Status: string;
  end;

  TFrmRappro = class(TForm)
  private
    FConn: TADOConnection;
    FQry: TADOQuery;
    FOrd: array of TQtyRow;     { quantités commandées }
    FRec: array of TQtyRow;     { quantités reçues (livraisons liées, non annulées) }
    FCmd: array of TCmdRow;
    FArtCodes, FArtDes: TStringList;
    FView: Integer;             { nombre de commandes affichées }
    FDetCount: Integer;
    grdCmd: TStringGrid;
    grdDet: TStringGrid;
    grdLiv: TStringGrid;
    chkOpen: TCheckBox;
    lblInfo: TLabel;
    lblTot: TLabel;

    function AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
    function AddButton(const ACaption: string; ALeft, ATop, AWidth: Integer;
      AHandler: TNotifyEvent): TButton;
    function MakeGrid(ALeft, ATop, AWidth, AHeight: Integer;
      const AHeads: array of string; const AWidths: array of Integer): TStringGrid;
    procedure BuildUI;
    procedure LoadData;
    function OrdQty(const ABc, AArt: string): Integer;
    function RecQty(const ABc, AArt: string): Integer;
    function CmdStatus(const ABc: string): string;
    function LineStatus(AOrd, ARec: Integer): string;
    function Designation(const ACode: string): string;
    procedure ApplyFilter(const AKeepCode: string);
    procedure ShowDetail(ARow: Integer);
    procedure ClearGrid(AGrid: TStringGrid);
    procedure PrintRappro;

    procedure FilterClick(Sender: TObject);
    procedure GridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure PrintClick(Sender: TObject);
    procedure CloseClick(Sender: TObject);
  public
    constructor CreateRappro(AOwner: TComponent; AConn: TADOConnection);
    destructor Destroy; override;
  end;

procedure OpenRappro(AConn: TADOConnection);

implementation

const
  ST_NONE = 'Non livrée';
  ST_PART = 'Partielle';
  ST_DONE = 'Complète';
  ST_EXC = 'Complète (excédent)';

procedure OpenRappro(AConn: TADOConnection);
var
  F: TFrmRappro;
begin
  F := TFrmRappro.CreateRappro(Application, AConn);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

{ TFrmRappro }

constructor TFrmRappro.CreateRappro(AOwner: TComponent; AConn: TADOConnection);
begin
  inherited CreateNew(AOwner);
  FConn := AConn;
  FQry := TADOQuery.Create(Self);
  FQry.Connection := FConn;
  FArtCodes := TStringList.Create;
  FArtDes := TStringList.Create;
  BuildUI;
  LoadData;
  ApplyFilter('');
end;

destructor TFrmRappro.Destroy;
begin
  FArtCodes.Free;
  FArtDes.Free;
  inherited;
end;

function TFrmRappro.AddLabel(const ACaption: string; ALeft, ATop: Integer): TLabel;
begin
  Result := TLabel.Create(Self);
  Result.Parent := Self;
  Result.Left := ALeft;
  Result.Top := ATop;
  Result.Caption := ACaption;
end;

function TFrmRappro.AddButton(const ACaption: string; ALeft, ATop,
  AWidth: Integer; AHandler: TNotifyEvent): TButton;
begin
  Result := TButton.Create(Self);
  Result.Parent := Self;
  Result.SetBounds(ALeft, ATop, AWidth, 28);
  Result.Caption := ACaption;
  Result.OnClick := AHandler;
end;

function TFrmRappro.MakeGrid(ALeft, ATop, AWidth, AHeight: Integer;
  const AHeads: array of string; const AWidths: array of Integer): TStringGrid;
var
  I: Integer;
begin
  Result := TStringGrid.Create(Self);
  Result.Parent := Self;
  Result.SetBounds(ALeft, ATop, AWidth, AHeight);
  Result.FixedCols := 0;
  Result.FixedRows := 1;
  Result.RowCount := 2;
  Result.ColCount := Length(AHeads);
  for I := 0 to High(AHeads) do
  begin
    Result.ColWidths[I] := AWidths[I];
    Result.Cells[I, 0] := AHeads[I];
  end;
  Result.Options := Result.Options + [goRowSelect] - [goEditing];
end;

procedure TFrmRappro.BuildUI;
begin
  BorderStyle := bsDialog;
  Position := poOwnerFormCenter;
  ClientWidth := 780;
  ClientHeight := 640;
  Caption := 'Rapprochement commande / livraison';

  AddLabel('Bons de commande (non annulés)', 16, 12);
  chkOpen := TCheckBox.Create(Self);
  chkOpen.Parent := Self;
  chkOpen.SetBounds(400, 10, 360, 17);
  chkOpen.Caption := 'Commandes non soldées seulement';
  chkOpen.OnClick := FilterClick;

  grdCmd := MakeGrid(16, 30, 748, 170, ['N° commande', 'Date', 'Service', 'Statut'],
    [130, 90, 330, 194]);
  grdCmd.OnSelectCell := GridSelectCell;
  lblInfo := AddLabel('', 16, 204);

  AddLabel('Détail par article de la commande sélectionnée', 16, 228);
  grdDet := MakeGrid(16, 246, 748, 190,
    ['Code', 'Désignation', 'Commandé', 'Reçu', 'Reste', 'Statut'],
    [90, 280, 80, 80, 80, 134]);
  lblTot := AddLabel('', 16, 440);

  AddLabel('Livraisons rattachées à cette commande', 16, 466);
  grdLiv := MakeGrid(16, 484, 748, 100,
    ['N° livraison', 'Date', 'Fournisseur', 'Statut'], [130, 90, 330, 194]);

  AddButton('Imprimer', 534, 600, 110, PrintClick);
  AddButton('Fermer', 654, 600, 110, CloseClick);
end;

procedure TFrmRappro.ClearGrid(AGrid: TStringGrid);
var
  R, C: Integer;
begin
  for R := 1 to AGrid.RowCount - 1 do
    for C := 0 to AGrid.ColCount - 1 do
      AGrid.Cells[C, R] := '';
  AGrid.RowCount := 2;
end;

procedure TFrmRappro.LoadData;
var
  N: Integer;
begin
  { articles }
  FArtCodes.Clear;
  FArtDes.Clear;
  FQry.Close;
  FQry.SQL.Text := 'SELECT Code_art, Designation FROM ARTICLE';
  FQry.Open;
  while not FQry.Eof do
  begin
    FArtCodes.Add(FQry.FieldByName('Code_art').AsString);
    FArtDes.Add(FQry.FieldByName('Designation').AsString);
    FQry.Next;
  end;
  FQry.Close;

  { quantités commandées (commandes non annulées) }
  SetLength(FOrd, 0);
  FQry.SQL.Text :=
    'SELECT c.Code_BC, c.Code_art, c.Qtt_cm FROM COMMANDER AS c ' +
    'INNER JOIN BON_DE_COMMANDE AS b ON c.Code_BC = b.Code_BC ' +
    'WHERE b.Annule IS NULL OR b.Annule = False';
  FQry.Open;
  while not FQry.Eof do
  begin
    N := Length(FOrd);
    SetLength(FOrd, N + 1);
    FOrd[N].Bc := FQry.Fields[0].AsString;
    FOrd[N].Art := FQry.Fields[1].AsString;
    FOrd[N].Qty := FQry.Fields[2].AsInteger;
    FQry.Next;
  end;
  FQry.Close;

  { quantités reçues : livraisons liées et non annulées }
  SetLength(FRec, 0);
  FQry.SQL.Text :=
    'SELECT l.Code_BC, c.Code_art, SUM(c.Qtt_liv) AS Qtt ' +
    'FROM CONTIENT_LIVRAISON AS c INNER JOIN BON_DE_LIVRAISON AS l ' +
    'ON c.Code_Bliv = l.Code_Bliv ' +
    'WHERE l.Code_BC IS NOT NULL AND (l.Annule IS NULL OR l.Annule = False) ' +
    'GROUP BY l.Code_BC, c.Code_art';
  FQry.Open;
  while not FQry.Eof do
  begin
    N := Length(FRec);
    SetLength(FRec, N + 1);
    FRec[N].Bc := FQry.Fields[0].AsString;
    FRec[N].Art := FQry.Fields[1].AsString;
    FRec[N].Qty := FQry.Fields[2].AsInteger;
    FQry.Next;
  end;
  FQry.Close;

  { commandes }
  SetLength(FCmd, 0);
  FQry.SQL.Text :=
    'SELECT b.Code_BC, b.Date_BC, s.Nom_serv FROM BON_DE_COMMANDE AS b ' +
    'LEFT JOIN SERVICE_NAFTAL AS s ON b.Code_serv = s.Code_serv ' +
    'WHERE b.Annule IS NULL OR b.Annule = False ' +
    'ORDER BY b.Date_BC DESC, b.Code_BC DESC';
  FQry.Open;
  while not FQry.Eof do
  begin
    N := Length(FCmd);
    SetLength(FCmd, N + 1);
    FCmd[N].Code := FQry.Fields[0].AsString;
    if FQry.Fields[1].IsNull then
      FCmd[N].DateTxt := ''
    else
      FCmd[N].DateTxt := FormatDateTime('dd/mm/yyyy', FQry.Fields[1].AsDateTime);
    FCmd[N].Service := FQry.Fields[2].AsString;
    FCmd[N].Status := CmdStatus(FCmd[N].Code);
    FQry.Next;
  end;
  FQry.Close;
end;

function TFrmRappro.OrdQty(const ABc, AArt: string): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to High(FOrd) do
    if (FOrd[I].Bc = ABc) and SameText(FOrd[I].Art, AArt) then
      Inc(Result, FOrd[I].Qty);
end;

function TFrmRappro.RecQty(const ABc, AArt: string): Integer;
var
  I: Integer;
begin
  Result := 0;
  for I := 0 to High(FRec) do
    if (FRec[I].Bc = ABc) and SameText(FRec[I].Art, AArt) then
      Inc(Result, FRec[I].Qty);
end;

function TFrmRappro.LineStatus(AOrd, ARec: Integer): string;
begin
  if AOrd = 0 then
    Result := 'Hors commande'
  else if ARec = 0 then
    Result := 'Non reçu'
  else if ARec < AOrd then
    Result := 'Partiel'
  else if ARec = AOrd then
    Result := 'Complet'
  else
    Result := 'Excédent';
end;

function TFrmRappro.CmdStatus(const ABc: string): string;
var
  I, R: Integer;
  AnyRec, AllDone, Excess: Boolean;
begin
  AnyRec := False;
  AllDone := True;
  Excess := False;
  for I := 0 to High(FOrd) do
    if FOrd[I].Bc = ABc then
    begin
      R := RecQty(ABc, FOrd[I].Art);
      if R > 0 then
        AnyRec := True;
      if R < OrdQty(ABc, FOrd[I].Art) then
        AllDone := False;
      if R > OrdQty(ABc, FOrd[I].Art) then
        Excess := True;
    end;
  for I := 0 to High(FRec) do
    if FRec[I].Bc = ABc then
    begin
      if FRec[I].Qty > 0 then
        AnyRec := True;
      if OrdQty(ABc, FRec[I].Art) = 0 then
        Excess := True;
    end;
  if not AnyRec then
    Result := ST_NONE
  else if not AllDone then
    Result := ST_PART
  else if Excess then
    Result := ST_EXC
  else
    Result := ST_DONE;
end;

function TFrmRappro.Designation(const ACode: string): string;
var
  I: Integer;
begin
  Result := '';
  for I := 0 to FArtCodes.Count - 1 do
    if SameText(FArtCodes[I], ACode) then
    begin
      Result := FArtDes[I];
      Exit;
    end;
end;

procedure TFrmRappro.ApplyFilter(const AKeepCode: string);
var
  I, R, SelRow, NOpen: Integer;
begin
  ClearGrid(grdCmd);
  R := 0;
  SelRow := 1;
  NOpen := 0;
  for I := 0 to High(FCmd) do
  begin
    if (FCmd[I].Status = ST_NONE) or (FCmd[I].Status = ST_PART) then
      Inc(NOpen);
    if chkOpen.Checked and not ((FCmd[I].Status = ST_NONE) or
       (FCmd[I].Status = ST_PART)) then
      Continue;
    Inc(R);
    if grdCmd.RowCount < R + 1 then
      grdCmd.RowCount := R + 1;
    grdCmd.Cells[0, R] := FCmd[I].Code;
    grdCmd.Cells[1, R] := FCmd[I].DateTxt;
    grdCmd.Cells[2, R] := FCmd[I].Service;
    grdCmd.Cells[3, R] := FCmd[I].Status;
    if (AKeepCode <> '') and SameText(FCmd[I].Code, AKeepCode) then
      SelRow := R;
  end;
  FView := R;
  if FView > 0 then
    grdCmd.RowCount := FView + 1;
  grdCmd.Row := SelRow;
  lblInfo.Caption := Format('%d commande(s) au total - %d non soldée(s). Affichées : %d.',
    [Length(FCmd), NOpen, FView]);
  ShowDetail(SelRow);
end;

procedure TFrmRappro.ShowDetail(ARow: Integer);
var
  Bc: string;
  I, R, O, Rc, SumO, SumR: Integer;
  Arts: TStringList;
begin
  ClearGrid(grdDet);
  ClearGrid(grdLiv);
  FDetCount := 0;
  lblTot.Caption := '';
  if (ARow < 1) or (ARow > FView) then
    Exit;
  Bc := grdCmd.Cells[0, ARow];

  Arts := TStringList.Create;
  try
    for I := 0 to High(FOrd) do
      if (FOrd[I].Bc = Bc) and (Arts.IndexOf(FOrd[I].Art) < 0) then
        Arts.Add(FOrd[I].Art);
    for I := 0 to High(FRec) do
      if (FRec[I].Bc = Bc) and (Arts.IndexOf(FRec[I].Art) < 0) then
        Arts.Add(FRec[I].Art);

    SumO := 0;
    SumR := 0;
    for R := 0 to Arts.Count - 1 do
    begin
      O := OrdQty(Bc, Arts[R]);
      Rc := RecQty(Bc, Arts[R]);
      Inc(SumO, O);
      Inc(SumR, Rc);
      if grdDet.RowCount < R + 2 then
        grdDet.RowCount := R + 2;
      grdDet.Cells[0, R + 1] := Arts[R];
      grdDet.Cells[1, R + 1] := Designation(Arts[R]);
      grdDet.Cells[2, R + 1] := IntToStr(O);
      grdDet.Cells[3, R + 1] := IntToStr(Rc);
      if O > Rc then
        grdDet.Cells[4, R + 1] := IntToStr(O - Rc)
      else
        grdDet.Cells[4, R + 1] := '0';
      grdDet.Cells[5, R + 1] := LineStatus(O, Rc);
    end;
    FDetCount := Arts.Count;
  finally
    Arts.Free;
  end;
  if FDetCount > 0 then
    grdDet.RowCount := FDetCount + 1;
  lblTot.Caption := Format('Total commandé : %d   Total reçu : %d', [SumO, SumR]);

  { livraisons rattachées }
  FQry.Close;
  FQry.SQL.Text :=
    'SELECT l.Code_Bliv, l.Date_liv, f.Nom, l.Annule ' +
    'FROM BON_DE_LIVRAISON AS l LEFT JOIN FOURNISSEUR AS f ' +
    'ON l.Code_fourn = f.Code_fourn WHERE l.Code_BC = :b ' +
    'ORDER BY l.Date_liv, l.Code_Bliv';
  FQry.Parameters.ParamByName('b').Value := Bc;
  FQry.Open;
  R := 0;
  while not FQry.Eof do
  begin
    Inc(R);
    if grdLiv.RowCount < R + 1 then
      grdLiv.RowCount := R + 1;
    grdLiv.Cells[0, R] := FQry.Fields[0].AsString;
    if FQry.Fields[1].IsNull then
      grdLiv.Cells[1, R] := ''
    else
      grdLiv.Cells[1, R] := FormatDateTime('dd/mm/yyyy', FQry.Fields[1].AsDateTime);
    grdLiv.Cells[2, R] := FQry.Fields[2].AsString;
    if FQry.Fields[3].AsBoolean then
      grdLiv.Cells[3, R] := 'Annulée (non comptée)'
    else
      grdLiv.Cells[3, R] := '';
    FQry.Next;
  end;
  FQry.Close;
end;

procedure TFrmRappro.FilterClick(Sender: TObject);
begin
  ApplyFilter('');
end;

procedure TFrmRappro.GridSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  ShowDetail(ARow);
end;

procedure TFrmRappro.CloseClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmRappro.PrintClick(Sender: TObject);
begin
  if (grdCmd.Row < 1) or (grdCmd.Row > FView) then
    raise Exception.Create('Sélectionnez une commande à imprimer.');
  if Printer.Printers.Count = 0 then
    raise Exception.Create('Aucune imprimante n''est installée.');
  with TPrintDialog.Create(Self) do
  try
    if not Execute then
      Exit;
  finally
    Free;
  end;
  PrintRappro;
end;

procedure TFrmRappro.PrintRappro;
const
  NCOL = 6;
  Pct: array[0..NCOL - 1] of Integer = (13, 37, 12, 12, 12, 14);
var
  Cv: TCanvas;
  Margin, Y, LineH, PageW, TableW, I, K, Used, Row: Integer;
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

  procedure NeedSpace(ALines: Integer);
  begin
    if Y + LineH * ALines > Printer.PageHeight - Margin then
    begin
      Printer.NewPage;
      Y := Margin;
    end;
  end;

begin
  Row := grdCmd.Row;
  Cv := Printer.Canvas;
  Printer.Title := 'Rapprochement ' + grdCmd.Cells[0, Row];
  Printer.BeginDoc;
  try
    Cv.Font.Name := 'Arial';
    Cv.Font.Size := 10;
    Margin := Printer.PageWidth div 14;
    PageW := Printer.PageWidth;
    TableW := PageW - 2 * Margin;
    LineH := Cv.TextHeight('Ag') * 3 div 2;
    Y := Margin;

    Heads[0] := 'Code'; Heads[1] := 'Désignation'; Heads[2] := 'Commandé';
    Heads[3] := 'Reçu'; Heads[4] := 'Reste'; Heads[5] := 'Statut';
    for K := 0 to NCOL - 1 do
      Right[K] := (K >= 2) and (K <= 4);
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
    Cv.TextOut(Margin, Y, 'RAPPROCHEMENT COMMANDE / LIVRAISON');
    Inc(Y, LineH * 2);
    Cv.Font.Size := 10;
    Cv.TextOut(Margin, Y, 'Commande n° ' + grdCmd.Cells[0, Row] + ' du ' +
      grdCmd.Cells[1, Row]);
    Inc(Y, LineH);
    Cv.Font.Style := [];
    Cv.TextOut(Margin, Y, Clip('Service : ' + grdCmd.Cells[2, Row], TableW));
    Inc(Y, LineH);
    Cv.TextOut(Margin, Y, 'Statut : ' + grdCmd.Cells[3, Row] +
      '   (situation au ' + FormatDateTime('dd/mm/yyyy', Now) + ')');
    Inc(Y, LineH * 2);

    TableHeader;
    for I := 1 to FDetCount do
    begin
      if Y + LineH * 2 > Printer.PageHeight - Margin then
      begin
        Printer.NewPage;
        Y := Margin;
        TableHeader;
      end;
      for K := 0 to NCOL - 1 do
        Cell(K, grdDet.Cells[K, I]);
      Inc(Y, LineH);
    end;
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    Inc(Y, LineH div 2);
    Cv.Font.Style := [fsBold];
    Cv.TextOut(Margin, Y, lblTot.Caption);
    Inc(Y, LineH * 2);

    NeedSpace(3);
    Cv.TextOut(Margin, Y, 'Livraisons rattachées :');
    Inc(Y, LineH);
    Cv.Font.Style := [];
    if (grdLiv.RowCount < 2) or (grdLiv.Cells[0, 1] = '') then
    begin
      Cv.TextOut(Margin, Y, 'Aucune.');
    end
    else
      for I := 1 to grdLiv.RowCount - 1 do
      begin
        NeedSpace(2);
        Cv.TextOut(Margin, Y, Clip(Trim(grdLiv.Cells[0, I] + '   ' +
          grdLiv.Cells[1, I] + '   ' + grdLiv.Cells[2, I] + '   ' +
          grdLiv.Cells[3, I]), TableW));
        Inc(Y, LineH);
      end;

    Printer.EndDoc;
  except
    Printer.Abort;
    raise;
  end;
end;

end.
