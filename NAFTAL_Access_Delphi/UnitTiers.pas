unit UnitTiers;

{ Gestion des fournisseurs et des services : liste, ajout, modification,
  suppression. Une seule fiche, construite en code, pour les deux tables. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Variants,
  Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.Grids, Vcl.Dialogs,
  Data.DB, Data.Win.ADODB;

type
  TTiersKind = (tkFournisseur, tkService);

  TFieldDef = record
    Caption: string;
    Col: string;
    MaxLen: Integer;
  end;

  TFrmTiers = class(TForm)
  private
    FConn: TADOConnection;
    FKind: TTiersKind;
    FQry: TADOQuery;
    FTable: string;
    FOrderCol: string;
    FFields: array of TFieldDef;
    FEdits: array of TEdit;
    FRowCount: Integer;
    FExisting: Boolean;
    grdList: TStringGrid;

    procedure DefineFields;
    procedure AddField(const ACaption, ACol: string; AMaxLen: Integer);
    procedure BuildUI;
    procedure LoadList;
    procedure ClearEdits;
    procedure ShowRecord(ARow: Integer);
    function KeyValue: string;
    function RecordExists(const AKey: string): Boolean;
    function ReferenceCount(const AKey: string): Integer;
    procedure SetParam(AIndex: Integer);

    procedure GridSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure NewClick(Sender: TObject);
    procedure SaveClick(Sender: TObject);
    procedure DeleteClick(Sender: TObject);
    procedure CloseClick(Sender: TObject);
  public
    constructor CreateTiers(AOwner: TComponent; AConn: TADOConnection;
      AKind: TTiersKind);
  end;

procedure OpenTiers(AConn: TADOConnection; AKind: TTiersKind);

implementation

procedure OpenTiers(AConn: TADOConnection; AKind: TTiersKind);
var
  F: TFrmTiers;
begin
  F := TFrmTiers.CreateTiers(Application, AConn, AKind);
  try
    F.ShowModal;
  finally
    F.Free;
  end;
end;

{ TFrmTiers }

constructor TFrmTiers.CreateTiers(AOwner: TComponent; AConn: TADOConnection;
  AKind: TTiersKind);
begin
  inherited CreateNew(AOwner);
  FConn := AConn;
  FKind := AKind;
  FQry := TADOQuery.Create(Self);
  FQry.Connection := FConn;
  DefineFields;
  BuildUI;
  LoadList;
  ClearEdits;
end;

procedure TFrmTiers.AddField(const ACaption, ACol: string; AMaxLen: Integer);
begin
  SetLength(FFields, Length(FFields) + 1);
  with FFields[High(FFields)] do
  begin
    Caption := ACaption;
    Col := ACol;
    MaxLen := AMaxLen;
  end;
end;

procedure TFrmTiers.DefineFields;
begin
  { Le 1er champ est la clé ; les 2 premiers sont obligatoires (NOT NULL). }
  if FKind = tkFournisseur then
  begin
    FTable := 'FOURNISSEUR';
    FOrderCol := 'Nom';
    AddField('Code', 'Code_fourn', 30);
    AddField('Nom', 'Nom', 80);
    AddField('Prénom', 'Prenom', 80);
    AddField('Adresse', 'Adresse', 255);
    AddField('Téléphone', 'TEL_f', 30);
    AddField('N° RC', 'Num_rc', 60);
    AddField('NIF', 'NIF', 60);
  end
  else
  begin
    FTable := 'SERVICE_NAFTAL';
    FOrderCol := 'Nom_serv';
    AddField('Code', 'Code_serv', 30);
    AddField('Nom du service', 'Nom_serv', 100);
    AddField('Téléphone', 'TEL', 30);
  end;
end;

procedure TFrmTiers.BuildUI;
var
  I, X, Y, W, Col, PerRow: Integer;
  Lbl: TLabel;
  Btn: TButton;

  function MakeButton(const ACaption: string; ALeft: Integer;
    AHandler: TNotifyEvent): TButton;
  begin
    Result := TButton.Create(Self);
    Result.Parent := Self;
    Result.SetBounds(ALeft, 0, 120, 28);
    Result.Caption := ACaption;
    Result.OnClick := AHandler;
  end;

begin
  BorderStyle := bsDialog;
  Position := poOwnerFormCenter;
  ClientWidth := 750;
  ClientHeight := 540;
  if FKind = tkFournisseur then
    Caption := 'Fournisseurs'
  else
    Caption := 'Services';

  { Champs de saisie : 3 par ligne }
  PerRow := 3;
  W := 220;
  SetLength(FEdits, Length(FFields));
  for I := 0 to High(FFields) do
  begin
    Col := I mod PerRow;
    X := 16 + Col * (W + 20);
    Y := 12 + (I div PerRow) * 50;
    Lbl := TLabel.Create(Self);
    Lbl.Parent := Self;
    Lbl.SetBounds(X, Y, W, 13);
    Lbl.Caption := FFields[I].Caption;
    FEdits[I] := TEdit.Create(Self);
    FEdits[I].Parent := Self;
    FEdits[I].SetBounds(X, Y + 18, W, 21);
    FEdits[I].MaxLength := FFields[I].MaxLen;
  end;

  Y := 12 + ((Length(FFields) + PerRow - 1) div PerRow) * 50;
  Btn := MakeButton('Nouveau', 16, NewClick);
  Btn.Top := Y;
  Btn := MakeButton('Enregistrer', 146, SaveClick);
  Btn.Top := Y;
  Btn := MakeButton('Supprimer', 276, DeleteClick);
  Btn.Top := Y;
  Btn := MakeButton('Fermer', 614, CloseClick);
  Btn.Top := Y;

  grdList := TStringGrid.Create(Self);
  grdList.Parent := Self;
  grdList.SetBounds(16, Y + 40, 718, ClientHeight - Y - 56);
  grdList.FixedCols := 0;
  grdList.FixedRows := 1;
  grdList.RowCount := 2;
  grdList.ColCount := Length(FFields);
  grdList.Options := grdList.Options + [goRowSelect] - [goEditing];
  grdList.DefaultColWidth := 718 div Length(FFields) - 1;
  for I := 0 to High(FFields) do
    grdList.Cells[I, 0] := FFields[I].Caption;
  grdList.OnSelectCell := GridSelectCell;
end;

procedure TFrmTiers.LoadList;
var
  I, R: Integer;
  Cols: string;
begin
  Cols := '';
  for I := 0 to High(FFields) do
  begin
    if I > 0 then
      Cols := Cols + ', ';
    Cols := Cols + FFields[I].Col;
  end;
  FQry.Close;
  FQry.SQL.Text := 'SELECT ' + Cols + ' FROM ' + FTable + ' ORDER BY ' + FOrderCol;
  FQry.Open;
  FRowCount := 0;
  R := 0;
  while not FQry.Eof do
  begin
    Inc(R);
    if grdList.RowCount < R + 1 then
      grdList.RowCount := R + 1;
    for I := 0 to High(FFields) do
      grdList.Cells[I, R] := FQry.Fields[I].AsString;
    Inc(FRowCount);
    FQry.Next;
  end;
  FQry.Close;
  if FRowCount = 0 then
  begin
    for I := 0 to High(FFields) do
      grdList.Cells[I, 1] := '';
    grdList.RowCount := 2;
  end
  else
    grdList.RowCount := FRowCount + 1;
end;

procedure TFrmTiers.ClearEdits;
var
  I: Integer;
begin
  for I := 0 to High(FEdits) do
    FEdits[I].Text := '';
  FEdits[0].ReadOnly := False;
  FExisting := False;
  if Showing then
    FEdits[0].SetFocus;
end;

procedure TFrmTiers.ShowRecord(ARow: Integer);
var
  I: Integer;
begin
  if (ARow < 1) or (ARow > FRowCount) then
    Exit;
  for I := 0 to High(FEdits) do
    FEdits[I].Text := grdList.Cells[I, ARow];
  FEdits[0].ReadOnly := True;
  FExisting := True;
end;

procedure TFrmTiers.GridSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
begin
  ShowRecord(ARow);
end;

function TFrmTiers.KeyValue: string;
begin
  Result := Trim(FEdits[0].Text);
end;

function TFrmTiers.RecordExists(const AKey: string): Boolean;
begin
  FQry.Close;
  FQry.SQL.Text := 'SELECT COUNT(*) AS N FROM ' + FTable +
    ' WHERE ' + FFields[0].Col + ' = :k';
  FQry.Parameters.ParamByName('k').Value := AKey;
  FQry.Open;
  Result := FQry.FieldByName('N').AsInteger > 0;
  FQry.Close;
end;

{ Nombre de bons qui référencent cet enregistrement. }
function TFrmTiers.ReferenceCount(const AKey: string): Integer;

  function CountIn(const ATable, ACol: string): Integer;
  begin
    FQry.Close;
    FQry.SQL.Text := 'SELECT COUNT(*) AS N FROM ' + ATable +
      ' WHERE ' + ACol + ' = :k';
    FQry.Parameters.ParamByName('k').Value := AKey;
    FQry.Open;
    Result := FQry.FieldByName('N').AsInteger;
    FQry.Close;
  end;

begin
  if FKind = tkFournisseur then
    Result := CountIn('BON_DE_LIVRAISON', 'Code_fourn')
  else
    Result := CountIn('BON_DE_COMMANDE', 'Code_serv') +
      CountIn('BON_D_AFFECTATION', 'Code_serv');
end;

procedure TFrmTiers.SetParam(AIndex: Integer);
var
  S: string;
  P: TParameter;
begin
  S := Trim(FEdits[AIndex].Text);
  P := FQry.Parameters.ParamByName('p' + IntToStr(AIndex));
  P.DataType := ftString;
  if S = '' then
    P.Value := Null
  else
    P.Value := S;
end;

procedure TFrmTiers.NewClick(Sender: TObject);
begin
  ClearEdits;
end;

procedure TFrmTiers.SaveClick(Sender: TObject);
var
  I: Integer;
  SQL, Cols, Vals: string;
begin
  if KeyValue = '' then
    raise Exception.Create('Saisissez le code.');
  if Trim(FEdits[1].Text) = '' then
    raise Exception.CreateFmt('Saisissez le champ « %s ».', [FFields[1].Caption]);

  if FExisting then
  begin
    SQL := '';
    for I := 1 to High(FFields) do
    begin
      if I > 1 then
        SQL := SQL + ', ';
      SQL := SQL + FFields[I].Col + ' = :p' + IntToStr(I);
    end;
    FQry.Close;
    FQry.SQL.Text := 'UPDATE ' + FTable + ' SET ' + SQL +
      ' WHERE ' + FFields[0].Col + ' = :p0';
  end
  else
  begin
    if RecordExists(KeyValue) then
      raise Exception.Create('Ce code existe déjà.');
    Cols := '';
    Vals := '';
    for I := 0 to High(FFields) do
    begin
      if I > 0 then
      begin
        Cols := Cols + ', ';
        Vals := Vals + ', ';
      end;
      Cols := Cols + FFields[I].Col;
      Vals := Vals + ':p' + IntToStr(I);
    end;
    FQry.Close;
    FQry.SQL.Text := 'INSERT INTO ' + FTable + ' (' + Cols + ') VALUES (' + Vals + ')';
  end;

  { Les paramètres se créent à l'affectation du SQL : les valeurs viennent après. }
  for I := 0 to High(FFields) do
    SetParam(I);
  FQry.ExecSQL;

  LoadList;
  ClearEdits;
  ShowMessage('Enregistrement effectué.');
end;

procedure TFrmTiers.DeleteClick(Sender: TObject);
var
  N: Integer;
begin
  if not FExisting then
    raise Exception.Create('Sélectionnez d''abord un enregistrement dans la liste.');
  N := ReferenceCount(KeyValue);
  if N > 0 then
    raise Exception.CreateFmt(
      'Suppression impossible : %d bon(s) utilisent cet enregistrement.', [N]);
  if MessageDlg('Supprimer « ' + Trim(FEdits[1].Text) + ' » ?',
    mtConfirmation, [mbYes, mbNo], 0) <> mrYes then
    Exit;

  FQry.Close;
  FQry.SQL.Text := 'DELETE FROM ' + FTable + ' WHERE ' + FFields[0].Col + ' = :k';
  FQry.Parameters.ParamByName('k').Value := KeyValue;
  FQry.ExecSQL;
  LoadList;
  ClearEdits;
end;

procedure TFrmTiers.CloseClick(Sender: TObject);
begin
  Close;
end;

end.
