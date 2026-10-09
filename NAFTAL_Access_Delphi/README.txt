APPLICATION DE GESTION DE STOCK - NAFTAL BOUIRA
Technologies : Microsoft Access (.accdb) + Delphi (VCL) + ADO

CONTENU
- 01_schema_access.sql : structure des tables Access (à exécuter/créer dans Access)
- StockNaftal.dpr : projet Delphi
- UnitMain.pas : interface Delphi et opérations principales via ADO
- UnitMain.dfm : formulaire principal
- GUIDE_INSTALLATION.txt : étapes de configuration

IMPORTANT
Ce paquet contient le code source de départ. Il ne contient pas un fichier .accdb déjà créé,
car la création d'un fichier Access natif requiert Microsoft Access ou le moteur ACE installé.
Créez une base vide StockNaftal.accdb, puis créez les tables avec les requêtes SQL fournies
(une requête CREATE TABLE à la fois dans Access > Créer > Création de requête > Mode SQL).
Les types et contraintes peuvent nécessiter de légères adaptations selon la version d'Access.

Configuration Delphi :
- Delphi avec composants VCL et ADO installés.
- Le projet utilise TADOConnection et TADOQuery.
- Modifier le chemin de StockNaftal.accdb dans UnitMain.pas si nécessaire.
- Fournisseur ACE conseillé : Microsoft.ACE.OLEDB.12.0 ou 16.0.
- La bitness de Delphi (32/64 bits) doit correspondre à celle du fournisseur ACE installé.

Fonctions de cette première base :
- afficher les articles
- ajouter un article
- enregistrer une entrée de stock
- enregistrer une sortie de stock avec contrôle de quantité disponible
- afficher le stock actualisé

Les bons détaillés, fournisseurs, services, impressions et authentification pourront être ajoutés
dans les prochaines versions à partir de ce socle. Les données d'essai ne sont pas des données réelles NAFTAL.
