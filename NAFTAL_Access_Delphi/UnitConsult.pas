unit UnitConsult;

{ Consultation, impression et annulation des bons déjà enregistrés. Liste des bons en haut, lignes du bon sélectionné en bas.
  L'impression passe par TPrinter (aucun composant d'état supplémentaire). }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Math,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.Grids, Vcl.Dialogs, Vcl.Printers,
  Vcl.Graphics, Data.DB, Data.Win.ADODB, UnitBon;

type
  TFrmConsult = class(TForm)
  private
    FConn: TADOConnection;
    FKind: TBonKind;
    FQry: TADOQuery;
    FBonCount: Integer;
    FLineCount: Integer;
    FAmount: Currency;
    FHeadCode, FHeadDate, FHeadTiers, FHeadCancel: string;
    FCancelInfo: TStringList;   { une entrée par bon : '' ou « Annulé le ... » }
    grdBons: TStringGrid;
    grdLines: TStringGrid;
    lblTotal: TLabel;
    lblEmpty: TLabel;
    lblCancel: TLabel;

    function AddButton(const ACaption: string; ALeft, ATop, AWidth: Integer;
      AHandler: TNotifyEvent): TButton;
    procedure BuildUI;
    procedure LoadBons;
    procedure LoadLines(ARow: Integer);
    function TotalQty: Integer;
    function TotalAmount: Currency;
    function FmtDate(AField: TField): string;

    procedure GridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure GridBonsDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure CancelClick(Sender: TObject);
    procedure CancelBon(const ACode, AMotif: string);
    procedure PrintClick(Sender: TObject);
    procedure CloseClick(Sender: TObject);
    procedure PrintBon;
  public
    constructor CreateConsult(AOwner: TComponent; AConn: TADOConnection;
      AKind: TBonKind);
    destructor Destroy; override;
  end;

procedure OpenConsult(AConn: TADOConnection; AKind: TBonKind);

implementation

procedure OpenConsult(AConn: TADOConnection; AKind: TBonKind);
var
  F: TFrmConsult;
begin
  F := TFrmConsult.CreateConsult(Application, AConn, AKind);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

{ TFrmConsult }

constructor TFrmConsult.CreateConsult(AOwner: TComponent;
  AConn: TADOConnection; AKind: TBonKind);
begin
  inherited CreateNew(AOwner);
  FConn := AConn;
  FKind := AKind;
  FQry := TADOQuery.Create(Self);
  FQry.Connection := FConn;
  FCancelInfo := TStringList.Create;
  BuildUI;
  LoadBons;
end;

destructor TFrmConsult.Destroy;
begin
  FCancelInfo.Free;
  inherited;
end;

function TFrmConsult.AddButton(const ACaption: string; ALeft, ATop,
  AWidth: Integer; AHandler: TNotifyEvent): TButton;
begin
  Result := TButton.Create(Self);
  Result.Parent := Self;
  Result.SetBounds(ALeft, ATop, AWidth, 28);
  Result.Caption := ACaption;
  Result.OnClick := AHandler;
end;

procedure TFrmConsult.BuildUI;
var
  L: TLabel;
  IsLiv: Boolean;
begin
  IsLiv := FKind = bkLivraison;
  BorderStyle := bsDialog;
  Position := poOwnerFormCenter;
  ClientWidth := 750;
  ClientHeight := 580;
  case FKind of
    bkLivraison: Caption := 'Consultation des bons de livraison';
    bkAffectation: Caption := 'Consultation des bons d''affectation';
  else
    Caption := 'Consultation des bons de commande';
  end;

  L := TLabel.Create(Self);
  L.Parent := Self;
  L.SetBounds(16, 10, 300, 13);
  L.Caption := 'Bons enregistrés (du plus récent au plus ancien)';

  grdBons := TStringGrid.Create(Self);
  grdBons.Parent := Self;
  grdBons.SetBounds(16, 28, 718, 200);
  grdBons.FixedCols := 0;
  grdBons.FixedRows := 1;
  grdBons.RowCount := 2;
  grdBons.ColCount := 4;
  grdBons.ColWidths[0] := 150;
  grdBons.ColWidths[1] := 110;
  grdBons.ColWidths[2] := 350;
  grdBons.ColWidths[3] := 90;
  grdBons.Options := grdBons.Options + [goRowSelect] - [goEditing];
  grdBons.Cells[0, 0] := 'N° du bon';
  grdBons.Cells[1, 0] := 'Date';
  if IsLiv then
    grdBons.Cells[2, 0] := 'Fournisseur'
  else
    grdBons.Cells[2, 0] := 'Service';
  grdBons.Cells[3, 0] := 'Statut';
  grdBons.OnSelectCell := GridSelectCell;
  grdBons.OnDrawCell := GridBonsDrawCell;

  lblEmpty := TLabel.Create(Self);
  lblEmpty.Parent := Self;
  lblEmpty.SetBounds(16, 232, 500, 13);
  lblEmpty.Caption := '';

  L := TLabel.Create(Self);
  L.Parent := Self;
  L.SetBounds(16, 252, 300, 13);
  L.Caption := 'Lignes du bon sélectionné';

  grdLines := TStringGrid.Create(Self);
  grdLines.Parent := Self;
  grdLines.SetBounds(16, 270, 718, 230);
  grdLines.FixedCols := 0;
  grdLines.FixedRows := 1;
  grdLines.RowCount := 2;
  grdLines.Options := grdLines.Options - [goEditing];
  if IsLiv then
  begin
    grdLines.ColCount := 5;
    grdLines.ColWidths[0] := 100;
    grdLines.ColWidths[1] := 295;
    grdLines.ColWidths[2] := 70;
    grdLines.ColWidths[3] := 110;
    grdLines.ColWidths[4] := 125;
    grdLines.Cells[3, 0] := 'Prix unitaire';
    grdLines.Cells[4, 0] := 'Montant';
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

  lblTotal := TLabel.Create(Self);
  lblTotal.Parent := Self;
  lblTotal.SetBounds(16, 510, 600, 13);
  lblTotal.Caption := '';

  lblCancel := TLabel.Create(Self);
  lblCancel.Parent := Self;
  lblCancel.SetBounds(16, 526, 718, 13);
  lblCancel.Font.Color := clRed;
  lblCancel.Font.Style := [fsBold];
  lblCancel.Caption := '';

  AddButton('Annuler le bon', 16, 544, 140, CancelClick);
  AddButton('Imprimer le bon', 480, 544, 140, PrintClick);
  AddButton('Fermer', 634, 544, 100, CloseClick);
end;

function TFrmConsult.FmtDate(AField: TField): string;
begin
  if AField.IsNull then
    Result := ''
  else
    Result := FormatDateTime('dd/mm/yyyy', AField.AsDateTime);
end;

procedure TFrmConsult.LoadBons;
var
  R: Integer;
begin
  FQry.Close;
  if FKind = bkLivraison then
    FQry.SQL.Text :=
      'SELECT b.Code_Bliv AS Code, b.Date_liv AS Dt, f.Nom AS Tiers, ' +
      'b.Annule AS Ann, b.Date_annul AS DtAnn, b.Motif_annul AS Motif ' +
      'FROM BON_DE_LIVRAISON AS b LEFT JOIN FOURNISSEUR AS f ' +
      'ON b.Code_fourn = f.Code_fourn ' +
      'ORDER BY b.Date_liv DESC, b.Code_Bliv DESC'
  else if FKind = bkCommande then
    FQry.SQL.Text :=
      'SELECT b.Code_BC AS Code, b.Date_BC AS Dt, s.Nom_serv AS Tiers, ' +
      'b.Annule AS Ann, b.Date_annul AS DtAnn, b.Motif_annul AS Motif ' +
      'FROM BON_DE_COMMANDE AS b LEFT JOIN SERVICE_NAFTAL AS s ' +
      'ON b.Code_serv = s.Code_serv ' +
      'ORDER BY b.Date_BC DESC, b.Code_BC DESC'
  else
    FQry.SQL.Text :=
      'SELECT b.Code_Aff AS Code, b.Date_Aff AS Dt, s.Nom_serv AS Tiers, ' +
      'b.Annule AS Ann, b.Date_annul AS DtAnn, b.Motif_annul AS Motif ' +
      'FROM BON_D_AFFECTATION AS b LEFT JOIN SERVICE_NAFTAL AS s ' +
      'ON b.Code_serv = s.Code_serv ' +
      'ORDER BY b.Date_Aff DESC, b.Code_Aff DESC';
  FQry.Open;
  FCancelInfo.Clear;
  R := 0;
  while not FQry.Eof do
  begin
    Inc(R);
    if grdBons.RowCount < R + 1 then
      grdBons.RowCount := R + 1;
    grdBons.Cells[0, R] := FQry.FieldByName('Code').AsString;
    grdBons.Cells[1, R] := FmtDate(FQry.FieldByName('Dt'));
    grdBons.Cells[2, R] := FQry.FieldByName('Tiers').AsString;
    if FQry.FieldByName('Ann').AsBoolean then
    begin
      grdBons.Cells[3, R] := 'Annulé';
      FCancelInfo.Add('Annulé le ' + FmtDate(FQry.FieldByName('DtAnn')) +
        ' - motif : ' + FQry.FieldByName('Motif').AsString);
    end
    else
    begin
      grdBons.Cells[3, R] := '';
      FCancelInfo.Add('');
    end;
    FQry.Next;
  end;
  FQry.Close;
  FBonCount := R;
  if FBonCount = 0 then
  begin
    grdBons.RowCount := 2;
    lblEmpty.Caption := 'Aucun bon enregistré.';
    LoadLines(0);
  end
  else
  begin
    grdBons.RowCount := FBonCount + 1;
    lblEmpty.Caption := Format('%d bon(s).', [FBonCount]);
    grdBons.Row := 1;
    LoadLines(1);
  end;
end;

procedure TFrmConsult.LoadLines(ARow: Integer);
var
  R, C: Integer;
  Q: Integer;
  P: Currency;
begin
  { vider l'ancien contenu }
  for R := 1 to grdLines.RowCount - 1 do
    for C := 0 to grdLines.ColCount - 1 do
      grdLines.Cells[C, R] := '';
  grdLines.RowCount := 2;
  FLineCount := 0;
  FAmount := 0;
  FHeadCode := '';
  FHeadDate := '';
  FHeadTiers := '';
  FHeadCancel := '';

  if (ARow >= 1) and (ARow <= FBonCount) then
  begin
    FHeadCancel := FCancelInfo[ARow - 1];
    FHeadCode := grdBons.Cells[0, ARow];
    FHeadDate := grdBons.Cells[1, ARow];
    FHeadTiers := grdBons.Cells[2, ARow];

    FQry.Close;
    if FKind = bkLivraison then
      FQry.SQL.Text :=
        'SELECT c.Code_art, a.Designation, c.Qtt_liv AS Qtt, c.Prix ' +
        'FROM CONTIENT_LIVRAISON AS c LEFT JOIN ARTICLE AS a ' +
        'ON c.Code_art = a.Code_art WHERE c.Code_Bliv = :b ORDER BY c.Code_art'
    else if FKind = bkCommande then
      FQry.SQL.Text :=
        'SELECT c.Code_art, a.Designation, c.Qtt_cm AS Qtt ' +
        'FROM COMMANDER AS c LEFT JOIN ARTICLE AS a ' +
        'ON c.Code_art = a.Code_art WHERE c.Code_BC = :b ORDER BY c.Code_art'
    else
      FQry.SQL.Text :=
        'SELECT c.Code_art, a.Designation, c.Qtt_aff AS Qtt ' +
        'FROM CONTIENT_AFFECTATION AS c LEFT JOIN ARTICLE AS a ' +
        'ON c.Code_art = a.Code_art WHERE c.Code_Aff = :b ORDER BY c.Code_art';
    FQry.Parameters.ParamByName('b').Value := FHeadCode;
    FQry.Open;
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
      begin
        P := FQry.FieldByName('Prix').AsCurrency;
        grdLines.Cells[3, R] := CurrToStrF(P, ffNumber, 2);
        grdLines.Cells[4, R] := CurrToStrF(Q * P, ffNumber, 2);
        FAmount := FAmount + Q * P;
      end;
      FQry.Next;
    end;
    FQry.Close;
    FLineCount := R;
    if FLineCount > 0 then
      grdLines.RowCount := FLineCount + 1;
  end;

  lblCancel.Caption := FHeadCancel;
  if FKind = bkLivraison then
    lblTotal.Caption := Format('Lignes : %d   Quantité totale : %d   Montant total : %s',
      [FLineCount, TotalQty, CurrToStrF(TotalAmount, ffNumber, 2)])
  else
    lblTotal.Caption := Format('Lignes : %d   Quantité totale : %d',
      [FLineCount, TotalQty]);
end;

function TFrmConsult.TotalQty: Integer;
var
  R: Integer;
begin
  Result := 0;
  for R := 1 to FLineCount do
    Inc(Result, StrToIntDef(grdLines.Cells[2, R], 0));
end;

function TFrmConsult.TotalAmount: Currency;
begin
  Result := FAmount;
end;

procedure TFrmConsult.GridSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  LoadLines(ARow);
end;

procedure TFrmConsult.GridBonsDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
begin
  if (ARow < 1) or (ARow > FBonCount) or (gdSelected in State) or
     (grdBons.Cells[3, ARow] = '') then
    Exit;
  with grdBons.Canvas do
  begin
    Brush.Color := RGB(220, 220, 220);
    FillRect(Rect);
    Font.Color := clGray;
    TextOut(Rect.Left + 4, Rect.Top + 2, grdBons.Cells[ACol, ARow]);
  end;
end;

{ Annule un bon : le bon est conservé et marqué « Annulé » (date, motif) ;
  le stock est corrigé dans la même transaction.
  - livraison   : le stock est diminué des quantités reçues (refusé si le
                  stock actuel est inférieur : des articles ont déjà été sortis) ;
  - affectation : le stock est rétabli ;
  - commande    : aucun effet sur le stock. }
procedure TFrmConsult.CancelBon(const ACode, AMotif: string);
var
  Tbl, Fld, Code: string;
  R, Q: Integer;
begin
  case FKind of
    bkLivraison:
      begin
        Tbl := 'BON_DE_LIVRAISON';
        Fld := 'Code_Bliv';
      end;
    bkAffectation:
      begin
        Tbl := 'BON_D_AFFECTATION';
        Fld := 'Code_Aff';
      end;
  else
    begin
      Tbl := 'BON_DE_COMMANDE';
      Fld := 'Code_BC';
    end;
  end;

  FConn.BeginTrans;
  try
    { Marquage d'abord : si un autre poste l'a déjà annulé, on s'arrête là. }
    FQry.Close;
    FQry.SQL.Text := 'UPDATE ' + Tbl +
      ' SET Annule = :a, Date_annul = :d, Motif_annul = :m' +
      ' WHERE ' + Fld + ' = :c AND (Annule IS NULL OR Annule = False)';
    FQry.Parameters.ParamByName('a').DataType := ftBoolean;
    FQry.Parameters.ParamByName('a').Value := True;
    FQry.Parameters.ParamByName('d').DataType := ftDateTime;
    FQry.Parameters.ParamByName('d').Value := Now;
    FQry.Parameters.ParamByName('m').Value := AMotif;
    FQry.Parameters.ParamByName('c').Value := ACode;
    FQry.ExecSQL;
    if FQry.RowsAffected = 0 then
      raise Exception.Create('Ce bon est déjà annulé (ou introuvable).');

    if FKind <> bkCommande then
      for R := 1 to FLineCount do
      begin
        Code := grdLines.Cells[0, R];
        Q := StrToInt(grdLines.Cells[2, R]);
        FQry.Close;
        if FKind = bkLivraison then
        begin
          FQry.SQL.Text :=
            'UPDATE ARTICLE SET Qtt_stock = Qtt_stock - :q ' +
            'WHERE Code_art = :c AND Qtt_stock >= :q2';
          FQry.Parameters.ParamByName('q').Value := Q;
          FQry.Parameters.ParamByName('c').Value := Code;
          FQry.Parameters.ParamByName('q2').Value := Q;
          FQry.ExecSQL;
          if FQry.RowsAffected = 0 then
            raise Exception.CreateFmt(
              'Annulation impossible : le stock de l''article %s est inférieur ' +
              'à la quantité livrée (%d). Des articles de ce bon ont déjà été sortis.',
              [Code, Q]);
        end
        else
        begin
          FQry.SQL.Text :=
            'UPDATE ARTICLE SET Qtt_stock = ' +
            'IIf(Qtt_stock IS NULL, 0, Qtt_stock) + :q WHERE Code_art = :c';
          FQry.Parameters.ParamByName('q').Value := Q;
          FQry.Parameters.ParamByName('c').Value := Code;
          FQry.ExecSQL;
          if FQry.RowsAffected = 0 then
            raise Exception.Create('Article introuvable : ' + Code);
        end;
      end;
    FConn.CommitTrans;
  except
    if FConn.InTransaction then
      FConn.RollbackTrans;
    raise;
  end;
end;

procedure TFrmConsult.CancelClick(Sender: TObject);
var
  Motif, Msg, Code: string;
  Row: Integer;
begin
  if FHeadCode = '' then
    raise Exception.Create('Sélectionnez un bon à annuler.');
  if FHeadCancel <> '' then
    raise Exception.Create('Ce bon est déjà annulé.');

  case FKind of
    bkLivraison:
      Msg := 'Le stock des articles sera DIMINUÉ des quantités livrées.';
    bkAffectation:
      Msg := 'Le stock des articles sera RÉTABLI (les quantités reviennent en stock).';
  else
    Msg := 'Une commande n''a aucun effet sur le stock.';
  end;
  if MessageDlg('Annuler le bon ' + FHeadCode + ' ?' + sLineBreak + Msg +
    sLineBreak + 'Cette opération est définitive.',
    mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
    Exit;

  Motif := '';
  if not InputQuery('Annulation du bon ' + FHeadCode,
    'Motif de l''annulation :', Motif) then
    Exit;
  Motif := Trim(Motif);
  if Motif = '' then
    raise Exception.Create('Le motif de l''annulation est obligatoire.');

  Code := FHeadCode;
  CancelBon(Code, Motif);

  LoadBons;
  for Row := 1 to FBonCount do
    if SameText(grdBons.Cells[0, Row], Code) then
    begin
      grdBons.Row := Row;
      LoadLines(Row);
      Break;
    end;
  ShowMessage('Bon ' + Code + ' annulé.');
end;

procedure TFrmConsult.CloseClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmConsult.PrintClick(Sender: TObject);
begin
  if FHeadCode = '' then
    raise Exception.Create('Sélectionnez un bon à imprimer.');
  if Printer.Printers.Count = 0 then
    raise Exception.Create('Aucune imprimante n''est installée.');
  with TPrintDialog.Create(Self) do
  try
    if not Execute then
      Exit;
  finally
    Free;
  end;
  PrintBon;
end;

procedure TFrmConsult.PrintBon;
var
  Cv: TCanvas;
  Margin, Y, LineH, PageW, R, I: Integer;
  ColX: array of Integer;     { bord gauche de chaque colonne }
  ColW: array of Integer;
  ColRight: array of Boolean; { colonne alignée à droite }
  Heads: array of string;
  IsLiv: Boolean;
  Title, TiersLabel: string;

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
    if ColRight[ACol] then
      Cv.TextOut(ColX[ACol] + ColW[ACol] - 4 - Cv.TextWidth(T), Y, T)
    else
      Cv.TextOut(ColX[ACol] + 4, Y, T);
  end;

  procedure TableHeader;
  var
    K: Integer;
  begin
    Cv.Font.Style := [fsBold];
    Cv.Pen.Width := 1;
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    for K := 0 to High(Heads) do
      Cell(K, Heads[K]);
    Inc(Y, LineH);
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    Cv.Font.Style := [];
    Inc(Y, 4);
  end;

begin
  IsLiv := FKind = bkLivraison;
  Cv := Printer.Canvas;
  Printer.Title := 'Bon ' + FHeadCode;
  Printer.BeginDoc;
  try
    Cv.Font.Name := 'Arial';
    Cv.Font.Size := 10;
    Margin := Printer.PageWidth div 14;
    PageW := Printer.PageWidth;
    LineH := Cv.TextHeight('Ag') * 3 div 2;
    Y := Margin;

    { colonnes : 5 pour une livraison, 3 pour une affectation }
    if IsLiv then
    begin
      SetLength(ColX, 5); SetLength(ColW, 5);
      SetLength(ColRight, 5); SetLength(Heads, 5);
      Heads[0] := 'Code'; Heads[1] := 'Désignation'; Heads[2] := 'Qté';
      Heads[3] := 'Prix unitaire'; Heads[4] := 'Montant';
      ColW[0] := (PageW - 2 * Margin) * 14 div 100;
      ColW[1] := (PageW - 2 * Margin) * 38 div 100;
      ColW[2] := (PageW - 2 * Margin) * 12 div 100;
      ColW[3] := (PageW - 2 * Margin) * 16 div 100;
      ColW[4] := (PageW - 2 * Margin) - ColW[0] - ColW[1] - ColW[2] - ColW[3];
      ColRight[0] := False; ColRight[1] := False; ColRight[2] := True;
      ColRight[3] := True; ColRight[4] := True;
    end
    else
    begin
      SetLength(ColX, 3); SetLength(ColW, 3);
      SetLength(ColRight, 3); SetLength(Heads, 3);
      Heads[0] := 'Code'; Heads[1] := 'Désignation'; Heads[2] := 'Qté';
      ColW[0] := (PageW - 2 * Margin) * 18 div 100;
      ColW[1] := (PageW - 2 * Margin) * 66 div 100;
      ColW[2] := (PageW - 2 * Margin) - ColW[0] - ColW[1];
      ColRight[0] := False; ColRight[1] := False; ColRight[2] := True;
    end;
    ColX[0] := Margin;
    for I := 1 to High(ColX) do
      ColX[I] := ColX[I - 1] + ColW[I - 1];

    { en-tête du document }
    Cv.Font.Style := [fsBold];
    Cv.Font.Size := 12;
    Cv.TextOut(Margin, Y, 'NAFTAL - District de Bouira');
    Inc(Y, LineH * 2);
    if IsLiv then
    begin
      Title := 'BON DE LIVRAISON N° ' + FHeadCode;
      TiersLabel := 'Fournisseur : ';
    end
    else if FKind = bkCommande then
    begin
      Title := 'BON DE COMMANDE N° ' + FHeadCode;
      TiersLabel := 'Service demandeur : ';
    end
    else
    begin
      Title := 'BON D''AFFECTATION N° ' + FHeadCode;
      TiersLabel := 'Service : ';
    end;
    Cv.Font.Size := 14;
    Cv.TextOut(Margin, Y, Title);
    Inc(Y, LineH * 2);
    if FHeadCancel <> '' then
    begin
      Cv.Font.Size := 11;
      Cv.TextOut(Margin, Y, Clip('*** BON ANNULÉ *** ' + FHeadCancel,
        PageW - 2 * Margin));
      Inc(Y, LineH * 2);
    end;
    Cv.Font.Size := 10;
    Cv.Font.Style := [];
    Cv.TextOut(Margin, Y, 'Date : ' + FHeadDate);
    Inc(Y, LineH);
    Cv.TextOut(Margin, Y, TiersLabel + FHeadTiers);
    Inc(Y, LineH * 2);

    TableHeader;
    for R := 1 to FLineCount do
    begin
      if Y + LineH * 6 > Printer.PageHeight - Margin then
      begin
        Printer.NewPage;
        Y := Margin;
        TableHeader;
      end;
      for I := 0 to High(Heads) do
        Cell(I, grdLines.Cells[I, R]);
      Inc(Y, LineH);
    end;

    { totaux et signatures }
    Cv.MoveTo(Margin, Y);
    Cv.LineTo(PageW - Margin, Y);
    Inc(Y, LineH div 2);
    Cv.Font.Style := [fsBold];
    Cv.TextOut(Margin, Y, Format('Quantité totale : %d', [TotalQty]));
    if IsLiv then
    begin
      Inc(Y, LineH);
      Cv.TextOut(Margin, Y, 'Montant total : ' + CurrToStrF(TotalAmount, ffNumber, 2));
    end;
    Cv.Font.Style := [];
    Inc(Y, LineH * 3);
    if Y + LineH * 5 > Printer.PageHeight - Margin then
    begin
      Printer.NewPage;
      Y := Margin;
    end;
    if FKind = bkCommande then
      Cv.TextOut(Margin, Y, 'Le demandeur')
    else
      Cv.TextOut(Margin, Y, 'Le magasinier');
    if IsLiv then
      Cv.TextOut(PageW - Margin - Cv.TextWidth('Le fournisseur'), Y, 'Le fournisseur')
    else if FKind = bkCommande then
      Cv.TextOut(PageW - Margin - Cv.TextWidth('Le responsable approvisionnement'), Y,
        'Le responsable approvisionnement')
    else
      Cv.TextOut(PageW - Margin - Cv.TextWidth('Le responsable du service'), Y,
        'Le responsable du service');

    Printer.EndDoc;
  except
    Printer.Abort;
    raise;
  end;
end;

end.
