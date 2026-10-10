unit UnitSchema;

{ Création et mise à niveau automatiques de la base Access.
  Au premier lancement, StockNaftal.accdb est créé (à côté de l'exécutable, ou
  dans le dossier LOCALAPPDATA\StockNaftal s'il n'est pas modifiable) avec toutes
  les tables. Aux lancements suivants, les tables ou colonnes manquantes
  (versions plus anciennes de la base) sont ajoutées : plus besoin de lancer les
  scripts 01_ / 03_ / 04_ à la main. Le contenu de la base n'est jamais effacé. }

interface

uses
  Winapi.Windows, System.SysUtils, System.Classes, System.Win.ComObj,
  Data.DB, Data.Win.ADODB;

const
  DB_FILE = 'StockNaftal.accdb';

{ Retourne le chemin de la base ; la crée (fichier vide) si elle n'existe pas. }
function EnsureDatabase(out ACreated: Boolean): string;

{ Crée les tables manquantes et ajoute les colonnes manquantes. }
procedure EnsureSchema(AConn: TADOConnection);

function ProviderList: TArray<string>;

implementation

type
  TTableDef = record
    Name: string;
    SQL: string;
  end;

  TColumnDef = record
    Table, Column, DDL: string;
  end;

const
  TABLES: array[0..8] of TTableDef = (
    (Name: 'ARTICLE'; SQL:
      'CREATE TABLE ARTICLE ( ' +
      'Code_art TEXT(30) NOT NULL, ' +
      'Designation TEXT(120) NOT NULL, ' +
      'Type_art TEXT(60), ' +
      'Qtt_stock LONG, ' +
      'Seuil_alerte LONG, ' +
      'CONSTRAINT PK_ARTICLE PRIMARY KEY (Code_art) ' +
      ')'),
    (Name: 'FOURNISSEUR'; SQL:
      'CREATE TABLE FOURNISSEUR ( ' +
      'Code_fourn TEXT(30) NOT NULL, ' +
      'Nom TEXT(80) NOT NULL, ' +
      'Prenom TEXT(80), ' +
      'Adresse TEXT(255), ' +
      'TEL_f TEXT(30), ' +
      'Num_rc TEXT(60), ' +
      'NIF TEXT(60), ' +
      'CONSTRAINT PK_FOURNISSEUR PRIMARY KEY (Code_fourn) ' +
      ')'),
    (Name: 'SERVICE_NAFTAL'; SQL:
      'CREATE TABLE SERVICE_NAFTAL ( ' +
      'Code_serv TEXT(30) NOT NULL, ' +
      'Nom_serv TEXT(100) NOT NULL, ' +
      'TEL TEXT(30), ' +
      'CONSTRAINT PK_SERVICE PRIMARY KEY (Code_serv) ' +
      ')'),
    (Name: 'BON_DE_COMMANDE'; SQL:
      'CREATE TABLE BON_DE_COMMANDE ( ' +
      'Code_BC TEXT(30) NOT NULL, ' +
      'Date_BC DATETIME, ' +
      'Code_serv TEXT(30), ' +
      'Annule YESNO, ' +
      'Date_annul DATETIME, ' +
      'Motif_annul TEXT(255), ' +
      'CONSTRAINT PK_BON_COMMANDE PRIMARY KEY (Code_BC) ' +
      ')'),
    (Name: 'COMMANDER'; SQL:
      'CREATE TABLE COMMANDER ( ' +
      'Code_art TEXT(30) NOT NULL, ' +
      'Code_BC TEXT(30) NOT NULL, ' +
      'Qtt_cm LONG, ' +
      'CONSTRAINT PK_COMMANDER PRIMARY KEY (Code_art, Code_BC) ' +
      ')'),
    (Name: 'BON_DE_LIVRAISON'; SQL:
      'CREATE TABLE BON_DE_LIVRAISON ( ' +
      'Code_Bliv TEXT(30) NOT NULL, ' +
      'Date_liv DATETIME, ' +
      'Code_fourn TEXT(30), ' +
      'Code_BC TEXT(30), ' +
      'Annule YESNO, ' +
      'Date_annul DATETIME, ' +
      'Motif_annul TEXT(255), ' +
      'CONSTRAINT PK_BON_LIVRAISON PRIMARY KEY (Code_Bliv) ' +
      ')'),
    (Name: 'CONTIENT_LIVRAISON'; SQL:
      'CREATE TABLE CONTIENT_LIVRAISON ( ' +
      'Code_art TEXT(30) NOT NULL, ' +
      'Code_Bliv TEXT(30) NOT NULL, ' +
      'Qtt_liv LONG, ' +
      'Prix CURRENCY, ' +
      'CONSTRAINT PK_CONTIENT_LIVRAISON PRIMARY KEY (Code_art, Code_Bliv) ' +
      ')'),
    (Name: 'BON_D_AFFECTATION'; SQL:
      'CREATE TABLE BON_D_AFFECTATION ( ' +
      'Code_Aff TEXT(30) NOT NULL, ' +
      'Date_Aff DATETIME, ' +
      'Code_serv TEXT(30), ' +
      'Annule YESNO, ' +
      'Date_annul DATETIME, ' +
      'Motif_annul TEXT(255), ' +
      'CONSTRAINT PK_BON_AFFECTATION PRIMARY KEY (Code_Aff) ' +
      ')'),
    (Name: 'CONTIENT_AFFECTATION'; SQL:
      'CREATE TABLE CONTIENT_AFFECTATION ( ' +
      'Code_art TEXT(30) NOT NULL, ' +
      'Code_Aff TEXT(30) NOT NULL, ' +
      'Qtt_aff LONG, ' +
      'CONSTRAINT PK_CONTIENT_AFFECTATION PRIMARY KEY (Code_art, Code_Aff) ' +
      ')')
  );

  { colonnes ajoutées après la première version }
  COLUMNS: array[0..9] of TColumnDef = (
    (Table: 'BON_DE_COMMANDE'; Column: 'Annule'; DDL: 'YESNO'),
    (Table: 'BON_DE_COMMANDE'; Column: 'Date_annul'; DDL: 'DATETIME'),
    (Table: 'BON_DE_COMMANDE'; Column: 'Motif_annul'; DDL: 'TEXT(255)'),
    (Table: 'BON_DE_LIVRAISON'; Column: 'Code_BC'; DDL: 'TEXT(30)'),
    (Table: 'BON_DE_LIVRAISON'; Column: 'Annule'; DDL: 'YESNO'),
    (Table: 'BON_DE_LIVRAISON'; Column: 'Date_annul'; DDL: 'DATETIME'),
    (Table: 'BON_DE_LIVRAISON'; Column: 'Motif_annul'; DDL: 'TEXT(255)'),
    (Table: 'BON_D_AFFECTATION'; Column: 'Annule'; DDL: 'YESNO'),
    (Table: 'BON_D_AFFECTATION'; Column: 'Date_annul'; DDL: 'DATETIME'),
    (Table: 'BON_D_AFFECTATION'; Column: 'Motif_annul'; DDL: 'TEXT(255)')
  );

function ProviderList: TArray<string>;
begin
  SetLength(Result, 2);
  Result[0] := 'Microsoft.ACE.OLEDB.12.0';
  Result[1] := 'Microsoft.ACE.OLEDB.16.0';
end;

function LocalDbDir: string;
begin
  Result := IncludeTrailingPathDelimiter(GetEnvironmentVariable('LOCALAPPDATA')) +
    'StockNaftal';
end;

procedure CreateFile(const AProvider, APath: string);
var
  Cat: OleVariant;
begin
  Cat := CreateOleObject('ADOX.Catalog');
  Cat.Create('Provider=' + AProvider + ';Data Source=' + APath + ';');
end;

function TryCreate(const APath: string): Boolean;
var
  P: string;
begin
  Result := False;
  for P in ProviderList do
    try
      ForceDirectories(ExtractFilePath(APath));
      CreateFile(P, APath);
      if FileExists(APath) then
      begin
        Result := True;
        Exit;
      end;
    except
      { fournisseur absent ou dossier non modifiable : on essaie la suite }
    end;
end;

function EnsureDatabase(out ACreated: Boolean): string;
var
  P1, P2: string;
begin
  ACreated := False;
  P1 := ExtractFilePath(ParamStr(0)) + DB_FILE;
  P2 := IncludeTrailingPathDelimiter(LocalDbDir) + DB_FILE;
  if FileExists(P1) then
  begin
    Result := P1;
    Exit;
  end;
  if FileExists(P2) then
  begin
    Result := P2;
    Exit;
  end;
  if TryCreate(P1) then
    Result := P1
  else if TryCreate(P2) then
    Result := P2
  else
    raise Exception.Create('Impossible de créer la base de données ' + DB_FILE +
      '.' + sLineBreak + 'Installez « Microsoft Access Database Engine » ' +
      '(redistribuable gratuit) dans la même architecture 32/64 bits que ' +
      'l''application, ou placez un fichier ' + DB_FILE + ' à côté du programme.');
  ACreated := True;
end;

function SqlOk(Q: TADOQuery; const S: string): Boolean;
begin
  try
    Q.Close;
    Q.SQL.Text := S;
    Q.Open;
    Q.Close;
    Result := True;
  except
    Result := False;
  end;
end;

procedure EnsureSchema(AConn: TADOConnection);
var
  Q: TADOQuery;
  I: Integer;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := AConn;
    for I := 0 to High(TABLES) do
      if not SqlOk(Q, 'SELECT * FROM ' + TABLES[I].Name + ' WHERE 1 = 0') then
      begin
        Q.Close;
        Q.SQL.Text := TABLES[I].SQL;
        Q.ExecSQL;
      end;
    for I := 0 to High(COLUMNS) do
      if not SqlOk(Q, 'SELECT ' + COLUMNS[I].Column + ' FROM ' +
        COLUMNS[I].Table + ' WHERE 1 = 0') then
      begin
        Q.Close;
        Q.SQL.Text := 'ALTER TABLE ' + COLUMNS[I].Table + ' ADD COLUMN ' +
          COLUMNS[I].Column + ' ' + COLUMNS[I].DDL;
        Q.ExecSQL;
      end;
  finally
    Q.Free;
  end;
end;

end.
