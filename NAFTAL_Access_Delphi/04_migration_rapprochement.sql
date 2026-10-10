-- MIGRATION : rapprochement commande / livraison (colonne Code_BC sur BON_DE_LIVRAISON).
-- À exécuter UNE SEULE FOIS si votre base a été créée avec une ancienne version de
-- 01_schema_access.sql. Une instruction, en mode SQL Access.
-- Les livraisons déjà enregistrées n'ont pas de commande liée (colonne vide) : elles
-- n'apparaissent pas dans le rapprochement.

ALTER TABLE BON_DE_LIVRAISON ADD COLUMN Code_BC TEXT(30);
