APPLICATION DE GESTION DE STOCK - NAFTAL BOUIRA
Technologies : Microsoft Access (.accdb) + Delphi (VCL) + ADO

FICHIERS
- 01_schema_access.sql        structure des tables (référence ; la base est créée automatiquement)
- 02_donnees_test.sql         données d'essai fictives (facultatif)
- 03_migration_annulation.sql / 04_migration_rapprochement.sql
                              référence seulement : la mise à niveau est automatique
- StockNaftal.dpr             projet Delphi
- UnitMain      écran principal : articles, alerte de seuil, accès aux modules
- UnitBon       saisie et correction des bons de livraison, d'affectation et de commande
- UnitConsult   consultation, impression, annulation et correction des bons
- UnitTiers     fournisseurs et services
- UnitStock     état du stock, seuils d'alerte, impression
- UnitMouv      mouvements d'un article (entrées, sorties, solde)
- UnitRappro    rapprochement commande / livraison
- UnitSchema    création automatique de la base Access et mise à niveau des tables
- GUIDE_INSTALLATION.txt      installation et utilisation
- PLAN_ESSAI.txt              scénario d'essai pas à pas

PRINCIPES
- Le stock ne change que par un bon de livraison (entrée) ou d'affectation (sortie).
- Chaque enregistrement (bon + lignes + stock) est une transaction : tout ou rien.
- Un bon n'est jamais supprimé : il est annulé (motif) ou remplacé (correction).
- Les données d'essai ne sont pas des données réelles NAFTAL.

Ce code n'a pas été compilé dans l'environnement de développement : voir PLAN_ESSAI.txt.
