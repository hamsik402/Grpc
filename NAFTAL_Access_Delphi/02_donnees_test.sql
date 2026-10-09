-- DONNÉES D'ESSAI (fictives, pas des données réelles NAFTAL).
-- À exécuter une instruction à la fois dans Access (Mode SQL), après 01_schema_access.sql.
-- Les écrans Fournisseurs et Services permettent aussi de les saisir ;
-- ces lignes servent à tester rapidement les bons.

INSERT INTO FOURNISSEUR (Code_fourn, Nom, Prenom, Adresse, TEL_f) VALUES ('F001', 'Fournisseur Test 1', '', 'Bouira', '0000000001');

INSERT INTO FOURNISSEUR (Code_fourn, Nom, Prenom, Adresse, TEL_f) VALUES ('F002', 'Fournisseur Test 2', '', 'Alger', '0000000002');

INSERT INTO SERVICE_NAFTAL (Code_serv, Nom_serv, TEL) VALUES ('S001', 'Service Informatique', '0000000011');

INSERT INTO SERVICE_NAFTAL (Code_serv, Nom_serv, TEL) VALUES ('S002', 'Service Comptabilité', '0000000012');
