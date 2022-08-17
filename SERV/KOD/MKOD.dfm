object FKOD: TFKOD
  Left = 1063
  Top = 147
  BorderWidth = 6
  Caption = #1050#1086#1076' '#1087#1072#1088#1072#1084#1077#1090#1088#1072
  ClientHeight = 664
  ClientWidth = 859
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object BCheck1: TBevel
    Left = 0
    Top = 610
    Width = 859
    Height = 6
    Align = alBottom
    Shape = bsSpacer
    ExplicitLeft = -277
    ExplicitTop = 559
    ExplicitWidth = 819
  end
  object Bevel2: TBevel
    Left = 0
    Top = 633
    Width = 859
    Height = 6
    Align = alBottom
    Shape = bsSpacer
    ExplicitTop = 541
    ExplicitWidth = 542
  end
  object Bevel3: TBevel
    Left = 0
    Top = 60
    Width = 859
    Height = 6
    Align = alTop
    Shape = bsSpacer
    ExplicitTop = 520
    ExplicitWidth = 791
  end
  object PControl: TPageControl
    Left = 0
    Top = 66
    Width = 859
    Height = 544
    ActivePage = TS_Tab
    Align = alClient
    Images = FMAIN.ImgLst16
    TabOrder = 0
    object TS_Tab: TTabSheet
      Caption = #1058#1072#1073#1083#1080#1094#1099
      ImageIndex = 20
      OnResize = TS_Resize
      object Bevel4: TBevel
        Left = 0
        Top = 509
        Width = 851
        Height = 6
        Align = alBottom
        Shape = bsSpacer
        ExplicitLeft = -1
        ExplicitTop = 416
        ExplicitWidth = 783
      end
      object Bevel5: TBevel
        Left = 0
        Top = 478
        Width = 851
        Height = 6
        Align = alBottom
        Shape = bsSpacer
        ExplicitTop = 520
        ExplicitWidth = 791
      end
      object GTab: TDBGrid
        Left = 0
        Top = 0
        Width = 851
        Height = 478
        Align = alClient
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
      end
      object NTab: TDBNavigator
        Left = 0
        Top = 484
        Width = 851
        Height = 25
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        Align = alBottom
        Hints.Strings = (
          #1044#1086#1073#1072#1074#1080#1090#1100' '#1079#1072#1087#1080#1089#1100
          #1059#1076#1072#1083#1080#1090#1100' '#1079#1072#1087#1080#1089#1100)
        Kind = dbnHorizontal
        ParentShowHint = False
        ShowHint = False
        TabOrder = 1
      end
    end
    object TS_Col: TTabSheet
      Caption = #1057#1090#1086#1083#1073#1094#1099
      ImageIndex = 24
      OnResize = TS_Resize
      object Bevel6: TBevel
        Left = 0
        Top = 509
        Width = 851
        Height = 6
        Align = alBottom
        Shape = bsSpacer
        ExplicitTop = 520
        ExplicitWidth = 791
      end
      object Bevel7: TBevel
        Left = 0
        Top = 478
        Width = 851
        Height = 6
        Align = alBottom
        Shape = bsSpacer
        ExplicitLeft = 3
        ExplicitTop = 342
        ExplicitWidth = 783
      end
      object NCol: TDBNavigator
        Left = 0
        Top = 484
        Width = 851
        Height = 25
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        Align = alBottom
        Kind = dbnHorizontal
        ParentShowHint = False
        ShowHint = False
        TabOrder = 0
      end
      object GCol: TDBGrid
        Left = 0
        Top = 0
        Width = 851
        Height = 478
        Align = alClient
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
      end
    end
    object TS_Row: TTabSheet
      Caption = #1057#1090#1088#1086#1082#1080
      ImageIndex = 22
      OnResize = TS_Resize
      object Bevel8: TBevel
        Left = 0
        Top = 509
        Width = 851
        Height = 6
        Align = alBottom
        Shape = bsSpacer
        ExplicitTop = 520
        ExplicitWidth = 791
      end
      object Bevel9: TBevel
        Left = 0
        Top = 478
        Width = 851
        Height = 6
        Align = alBottom
        Shape = bsSpacer
        ExplicitLeft = -1
        ExplicitTop = 358
        ExplicitWidth = 783
      end
      object NRow: TDBNavigator
        Left = 0
        Top = 484
        Width = 851
        Height = 25
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        Align = alBottom
        Kind = dbnHorizontal
        ParentShowHint = False
        ShowHint = False
        TabOrder = 0
      end
      object GRow: TDBGrid
        Left = 0
        Top = 0
        Width = 851
        Height = 478
        Align = alClient
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'MS Sans Serif'
        TitleFont.Style = []
      end
    end
  end
  object CBModify: TCheckBox
    Left = 0
    Top = 616
    Width = 859
    Height = 17
    Align = alBottom
    Caption = #1056#1072#1079#1088#1077#1096#1080#1090#1100' '#1088#1077#1076#1072#1082#1090#1080#1088#1086#1074#1072#1085#1080#1077' '#1089#1087#1080#1089#1082#1072
    TabOrder = 1
    OnClick = CBModifyClick
    OnKeyPress = CBModifyKeyPress
  end
  object PTab: TPanel
    Left = 0
    Top = 0
    Width = 859
    Height = 20
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 2
    object LTabCaption: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 62
      Height = 14
      Align = alLeft
      AutoSize = False
      Caption = #1058#1072#1073#1083#1080#1094#1072
      ExplicitHeight = 17
    end
    object LTab: TDBText
      AlignWithMargins = True
      Left = 71
      Top = 3
      Width = 42
      Height = 14
      Align = alLeft
      ExplicitHeight = 13
    end
    object ETab: TDBText
      AlignWithMargins = True
      Left = 119
      Top = 3
      Width = 737
      Height = 14
      Align = alClient
      ExplicitLeft = 71
      ExplicitWidth = 42
      ExplicitHeight = 13
    end
  end
  object PCol: TPanel
    Left = 0
    Top = 20
    Width = 859
    Height = 20
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 3
    object LColCaption: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 62
      Height = 14
      Align = alLeft
      AutoSize = False
      Caption = #1057#1090#1086#1083#1073#1077#1094
      ExplicitHeight = 17
    end
    object LCol: TDBText
      AlignWithMargins = True
      Left = 71
      Top = 3
      Width = 42
      Height = 14
      Align = alLeft
      ExplicitHeight = 13
    end
    object ECol: TDBText
      AlignWithMargins = True
      Left = 119
      Top = 3
      Width = 737
      Height = 14
      Align = alClient
      ExplicitLeft = 71
      ExplicitWidth = 42
      ExplicitHeight = 13
    end
  end
  object PRow: TPanel
    Left = 0
    Top = 40
    Width = 859
    Height = 20
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 4
    object LRowCaption: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 62
      Height = 14
      Align = alLeft
      AutoSize = False
      Caption = #1057#1090#1088#1086#1082#1072
      ExplicitHeight = 17
    end
    object LRow: TDBText
      AlignWithMargins = True
      Left = 71
      Top = 3
      Width = 42
      Height = 14
      Align = alLeft
      ExplicitHeight = 13
    end
    object ERow: TDBText
      AlignWithMargins = True
      Left = 119
      Top = 3
      Width = 737
      Height = 14
      Align = alClient
      ExplicitLeft = 71
      ExplicitWidth = 42
      ExplicitHeight = 13
    end
  end
  object PBottom: TPanel
    Left = 0
    Top = 639
    Width = 859
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 5
    object BtnOk: TBitBtn
      Left = 0
      Top = 0
      Width = 249
      Height = 25
      Action = AOk
      Align = alLeft
      Caption = #1054#1082
      TabOrder = 0
      ExplicitLeft = 1
      ExplicitTop = 1
      ExplicitHeight = 23
    end
    object BtnClose: TBitBtn
      Left = 249
      Top = 0
      Width = 361
      Height = 25
      Action = AClose
      Align = alClient
      Caption = #1047#1072#1082#1088#1099#1090#1100
      TabOrder = 1
      ExplicitLeft = 250
      ExplicitTop = 1
      ExplicitWidth = 359
      ExplicitHeight = 23
    end
    object BtnCancel: TBitBtn
      Left = 610
      Top = 0
      Width = 249
      Height = 25
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 2
      ExplicitLeft = 609
      ExplicitTop = 1
      ExplicitHeight = 23
    end
  end
  object DS_TAB: TDataSource
    DataSet = T_TAB
    Left = 376
    Top = 32
  end
  object DS_ROW: TDataSource
    DataSet = T_ROW
    Left = 408
    Top = 32
  end
  object DS_COL: TDataSource
    DataSet = T_COL
    Left = 440
    Top = 32
  end
  object T_TAB: TADOTable
    AfterScroll = T_TABAfterScroll
    Left = 376
  end
  object T_ROW: TADOTable
    OnNewRecord = T_ROWNewRecord
    Left = 408
  end
  object T_COL: TADOTable
    OnNewRecord = T_COLNewRecord
    Left = 440
  end
  object AList: TActionList
    Images = FMAIN.ImgLstSys
    Left = 472
    object AOk: TAction
      Caption = #1054#1082
      ImageIndex = 0
    end
    object AClose: TAction
      Caption = #1047#1072#1082#1088#1099#1090#1100
      ImageIndex = 0
    end
    object ACancel: TAction
      Caption = #1054#1090#1084#1077#1085#1072
      ImageIndex = 1
    end
  end
end
