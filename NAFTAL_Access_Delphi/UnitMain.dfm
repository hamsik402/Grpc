object FrmMain: TFrmMain
  Left = 0
  Top = 0
  Caption = 'Gestion de stock - NAFTAL Bouira'
  ClientHeight = 560
  ClientWidth = 940
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OnCreate = FormCreate
  object pnlTop: TPanel
    Left = 0
    Top = 0
    Width = 940
    Height = 64
    Align = alTop
    Caption = ''
    TabOrder = 0
    object lblTitle: TLabel
      Left = 20
      Top = 18
      Width = 450
      Height = 28
      Caption = 'NAFTAL BOUIRA - GESTION DE STOCK'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clNavy
      Font.Height = -20
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object pnlActions: TPanel
    Left = 0
    Top = 64
    Width = 940
    Height = 176
    Align = alTop
    TabOrder = 1
    object lblCode: TLabel
      Left = 16
      Top = 14
      Width = 32
      Height = 13
      Caption = 'Code'
    end
    object lblDesignation: TLabel
      Left = 180
      Top = 14
      Width = 65
      Height = 13
      Caption = 'Désignation'
    end
    object lblType: TLabel
      Left = 440
      Top = 14
      Width = 24
      Height = 13
      Caption = 'Type'
    end
    object lblQty: TLabel
      Left = 620
      Top = 14
      Width = 55
      Height = 13
      Caption = 'Quantité'
    end
    object edtCode: TEdit
      Left = 16
      Top = 34
      Width = 145
      Height = 21
      TabOrder = 0
    end
    object edtDesignation: TEdit
      Left = 180
      Top = 34
      Width = 240
      Height = 21
      TabOrder = 1
    end
    object edtType: TEdit
      Left = 440
      Top = 34
      Width = 160
      Height = 21
      TabOrder = 2
    end
    object edtQty: TEdit
      Left = 620
      Top = 34
      Width = 100
      Height = 21
      TabOrder = 3
    end
    object btnAdd: TButton
      Left = 16
      Top = 78
      Width = 130
      Height = 28
      Caption = 'Ajouter article'
      TabOrder = 4
      OnClick = btnAddClick
    end
    object btnEntry: TButton
      Left = 160
      Top = 78
      Width = 130
      Height = 28
      Caption = 'Entrée stock'
      TabOrder = 5
      OnClick = btnEntryClick
    end
    object btnExit: TButton
      Left = 304
      Top = 78
      Width = 130
      Height = 28
      Caption = 'Sortie stock'
      TabOrder = 6
      OnClick = btnExitClick
    end
    object btnRefresh: TButton
      Left = 448
      Top = 78
      Width = 130
      Height = 28
      Caption = 'Actualiser'
      TabOrder = 7
      OnClick = btnRefreshClick
    end
    object btnLivraison: TButton
      Left = 16
      Top = 112
      Width = 130
      Height = 28
      Caption = 'Bon de livraison'
      TabOrder = 8
      OnClick = btnLivraisonClick
    end
    object btnAffectation: TButton
      Left = 160
      Top = 112
      Width = 130
      Height = 28
      Caption = 'Bon d''affectation'
      TabOrder = 9
      OnClick = btnAffectationClick
    end
    object btnFournisseurs: TButton
      Left = 304
      Top = 112
      Width = 130
      Height = 28
      Caption = 'Fournisseurs'
      TabOrder = 10
      OnClick = btnFournisseursClick
    end
    object btnServices: TButton
      Left = 448
      Top = 112
      Width = 130
      Height = 28
      Caption = 'Services'
      TabOrder = 11
      OnClick = btnServicesClick
    end
    object btnConsLiv: TButton
      Left = 16
      Top = 146
      Width = 130
      Height = 28
      Caption = 'Liste des livraisons'
      TabOrder = 12
      OnClick = btnConsLivClick
    end
    object btnConsAff: TButton
      Left = 160
      Top = 146
      Width = 130
      Height = 28
      Caption = 'Liste des affectations'
      TabOrder = 13
      OnClick = btnConsAffClick
    end
  end
  object grdArticles: TStringGrid
    Left = 0
    Top = 204
    Width = 940
    Height = 356
    Align = alClient
    ColCount = 5
    DefaultColWidth = 170
    FixedCols = 0
    RowCount = 2
    TabOrder = 2
  end
  object ADOConnection1: TADOConnection
    LoginPrompt = False
    Left = 840
    Top = 16
  end
  object ADOQuery1: TADOQuery
    Connection = ADOConnection1
    Left = 880
    Top = 16
  end
end
