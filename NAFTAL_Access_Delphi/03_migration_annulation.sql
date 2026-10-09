-- MIGRATION : annulation des bons (colonnes Annule, Date_annul, Motif_annul).
-- À exécuter UNE SEULE FOIS si votre base StockNaftal.accdb a été créée avec une
-- ancienne version de 01_schema_access.sql. Une instruction à la fois, en mode SQL Access.
-- Si une instruction répond que la colonne existe déjà, passez à la suivante.

ALTER TABLE BON_DE_LIVRAISON ADD COLUMN Annule YESNO;
ALTER TABLE BON_DE_LIVRAISON ADD COLUMN Date_annul DATETIME;
ALTER TABLE BON_DE_LIVRAISON ADD COLUMN Motif_annul TEXT(255);

ALTER TABLE BON_D_AFFECTATION ADD COLUMN Annule YESNO;
ALTER TABLE BON_D_AFFECTATION ADD COLUMN Date_annul DATETIME;
ALTER TABLE BON_D_AFFECTATION ADD COLUMN Motif_annul TEXT(255);

ALTER TABLE BON_DE_COMMANDE ADD COLUMN Annule YESNO;
ALTER TABLE BON_DE_COMMANDE ADD COLUMN Date_annul DATETIME;
ALTER TABLE BON_DE_COMMANDE ADD COLUMN Motif_annul TEXT(255);
