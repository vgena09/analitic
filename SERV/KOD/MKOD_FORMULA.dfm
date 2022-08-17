object FKOD_FORMULA: TFKOD_FORMULA
  Left = 1060
  Top = 312
  BorderWidth = 6
  Caption = #1056#1077#1076#1072#1082#1090#1086#1088' '#1092#1086#1088#1084#1091#1083
  ClientHeight = 446
  ClientWidth = 808
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnHide = FormHide
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 411
    Width = 808
    Height = 10
    Align = alBottom
    Shape = bsTopLine
    ExplicitTop = 143
    ExplicitWidth = 626
  end
  object PBottom: TPanel
    Left = 0
    Top = 421
    Width = 808
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    ExplicitWidth = 679
    object Bevel3: TBevel
      Left = 604
      Top = 0
      Width = 16
      Height = 25
      Align = alRight
      Shape = bsSpacer
      ExplicitLeft = 416
    end
    object BtnOk: TBitBtn
      Left = 436
      Top = 0
      Width = 168
      Height = 25
      Action = AOk
      Align = alRight
      Caption = #1054#1082
      TabOrder = 0
      ExplicitLeft = 307
    end
    object BtnCancel: TBitBtn
      Left = 620
      Top = 0
      Width = 188
      Height = 25
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000FF00FF00FF00
        FF00FF00FF004B4EADFF696EFFFFFF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00363CF7FF3E40B0FFFF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF000000BFFF2928FFFF2828A6FF1E1EEAFFFF00FF00FF00FF00FF00FF00FF00
        FF00282CE9FF1716A5FF1A17FFFF0000C9FFFF00FF00FF00FF00FF00FF000E14
        C2FF0000D2FF0000DCFF0000F9FF1516A6FF282AE7FFFF00FF00FF00FF003335
        E5FF1817A6FF0000FCFF0000DEFF0B0ED9FF1819CCFFFF00FF00EFEFF9FF908F
        FAFF0000CAFF0000B9FF0000D4FF0000F0FF1010A6FF585BD0FF6267CFFF0F12
        A6FF0000F3FF0000D0FF0000BBFF0000CFFF5752ECFFE6E6F3FFFF00FF00FBFB
        FDFF0B0BBCFF0101CAFF0000B5FF0000C6FF0000D9FF3332CEFF3939CFFF0000
        DBFF0000C3FF0000B4FF0604D0FF0F13C0FFF5F5F9FFFF00FF00FF00FF00FF00
        FF00FF00FF001A16C8FF0001C1FF00009CFF0000AFFF0000C4FF0000C7FF0000
        AEFF0000A2FF0503C6FF1716C4FFFF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF001918BAFF1311BBFF000088FF00008DFF000090FF0000
        86FF0F0EB4FF1F1CBAFFFF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF0064619FFF1F1CC2FF000099FF000096FF000096FF0000
        96FF1715C2FF514F94FFFF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00363646FF0C0DC3FF0000B3FF00009BFF00009EFF00009EFF0000
        A1FF0000B5FF0302B6FF2A2A3CFFFF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF002B2B2BFF1312DEFF0000F8FF0000C0FF0000C1FF2925F0FF1919EBFF0000
        C2FF0000C4FF0000F6FF0000D8FF1E1E1EFFFF00FF00FF00FF00FF00FF002018
        0CFF3431D4FF0E0AFFFF0605EAFF100CEAFF2320FFFF736FEEFF5A56E6FF231F
        FFFF1511EAFF0A07EAFF0F0DFFFF1F1DE0FF0C0501FFFF00FF00DDDDDDFFC3AC
        F9FF9C8CFFFF4B40FFFF5D4BFCFF9180FFFF2D2FD6FFFF00FF00FF00FF003732
        D0FF8D7AFFFF5E4EF9FF5A4EFFFF8374FFFF6F58D2FFD5D5D5FFFF00FF00C7C7
        C7FF373399FFD8C9FFFFDFD1FFFF3F41D9FFFF00FF00FF00FF00FF00FF00FF00
        FF004946DCFFE0D0FFFFDDCEFFFF48439AFFC2C2C2FFFF00FF00FF00FF00FF00
        FF00D0D0D1FFE9E0FFFFB8B8FFFFFAFAFDFFFF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00A4A2FFFFBAB3FEFFCACACBFFFF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF009F9EDBFFBEBEEFFFFF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00C5C6F7FF9C9BD9FFFF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00}
      ModalResult = 2
      TabOrder = 1
      ExplicitLeft = 491
    end
  end
  object PFormula: TPanel
    Left = 0
    Top = 22
    Width = 808
    Height = 57
    Align = alTop
    AutoSize = True
    BevelOuter = bvNone
    TabOrder = 1
    ExplicitWidth = 679
    object LFormula: TLabel
      Left = 0
      Top = 10
      Width = 48
      Height = 13
      Align = alTop
      Caption = #1060#1086#1088#1084#1091#1083#1072
    end
    object Bevel2: TBevel
      Left = 0
      Top = 47
      Width = 808
      Height = 10
      Align = alTop
      Shape = bsSpacer
      ExplicitTop = 88
      ExplicitWidth = 626
    end
    object Bevel4: TBevel
      Left = 0
      Top = 0
      Width = 808
      Height = 10
      Align = alTop
      Shape = bsSpacer
      ExplicitTop = -3
      ExplicitWidth = 718
    end
    object CBFormula: TComboBox
      Left = 0
      Top = 23
      Width = 808
      Height = 24
      Align = alTop
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clBlue
      Font.Height = -13
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 0
      Text = 'CBFormula'
      OnChange = CBFormulaChange
      OnKeyDown = CBFormulaKeyDown
    end
  end
  object CoolBar1: TCoolBar
    Left = 0
    Top = 0
    Width = 808
    Height = 22
    AutoSize = True
    Bands = <
      item
        Control = ToolBar1
        ImageIndex = -1
        MinHeight = 22
        Width = 806
      end>
    EdgeBorders = []
    Images = FMAIN.ImgLstSys
    ExplicitWidth = 851
    object ToolBar1: TToolBar
      AlignWithMargins = True
      Left = 11
      Top = 0
      Width = 476
      Height = 22
      Align = alNone
      AutoSize = True
      ButtonWidth = 92
      Caption = 'ToolBar1'
      DockSite = True
      Images = FMAIN.ImgLstSys
      List = True
      ParentShowHint = False
      ShowCaptions = True
      ShowHint = True
      TabOrder = 0
      object BtnSave: TToolButton
        Left = 0
        Top = 0
        Action = ASave
      end
      object BtnOpen: TToolButton
        Left = 92
        Top = 0
        Action = AOpen
      end
      object BtnOpenStandart: TToolButton
        Left = 184
        Top = 0
        Action = AOpenStandart
      end
      object Separator1: TToolButton
        Left = 276
        Top = 0
        Width = 8
        Caption = 'Separator1'
        ImageIndex = 9
        Style = tbsSeparator
      end
      object BtnClear: TToolButton
        Left = 284
        Top = 0
        Action = AClear
      end
      object ToolButton1: TToolButton
        Left = 376
        Top = 0
        Width = 8
        Caption = 'ToolButton1'
        ImageIndex = 8
        Style = tbsSeparator
      end
      object BtnAddKod: TToolButton
        Left = 384
        Top = 0
        Action = AKod
      end
    end
  end
  object PDesc: TPanel
    Left = 0
    Top = 79
    Width = 808
    Height = 332
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 3
    OnResize = PDescResize
    ExplicitWidth = 679
    object LDesc: TLabel
      Left = 0
      Top = 0
      Width = 808
      Height = 13
      Align = alTop
      Caption = #1054#1087#1080#1089#1072#1085#1080#1077
      ExplicitWidth = 50
    end
    object EDesc: TRichEdit
      Left = 0
      Top = 13
      Width = 808
      Height = 319
      Align = alClient
      BevelInner = bvNone
      BevelOuter = bvNone
      Color = clBtnFace
      Font.Charset = RUSSIAN_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'MS Sans Serif'
      Font.Style = []
      Lines.Strings = (
        'EDesc')
      ParentFont = False
      TabOrder = 0
      ExplicitWidth = 679
    end
  end
  object ActionList1: TActionList
    Images = FMAIN.ImgLstSys
    Left = 224
    Top = 120
    object AOk: TAction
      Caption = #1054#1082
      ImageIndex = 0
      OnExecute = AOkExecute
    end
    object ASave: TAction
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' ...'
      ImageIndex = 3
      ShortCut = 113
      OnExecute = ASaveExecute
    end
    object AOpen: TAction
      Caption = #1054#1090#1082#1088#1099#1090#1100' ...'
      Hint = #1054#1090#1082#1088#1099#1090#1100' '#1088#1072#1085#1077#1077' '#1089#1086#1093#1088#1072#1085#1077#1085#1085#1091#1102' '#1092#1086#1088#1084#1091#1083#1091
      ImageIndex = 2
      ShortCut = 114
      OnExecute = AOpenExecute
    end
    object AOpenStandart: TAction
      Caption = #1054#1090#1082#1088#1099#1090#1100
      ImageIndex = 2
      ShortCut = 16498
      OnExecute = AOpenStandartExecute
    end
    object AClear: TAction
      Caption = #1054#1095#1080#1089#1090#1080#1090#1100
      ImageIndex = 4
      OnExecute = AClearExecute
    end
    object AKod: TAction
      Caption = #1050#1086#1076' ...'
      ImageIndex = 7
      ShortCut = 115
      OnExecute = AKodExecute
    end
  end
  object OpenDlg: TOpenDialog
    DefaultExt = 'set1'
    Filter = #1060#1072#1081#1083#1099' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080'|*.set1|'#1042#1089#1077' '#1092#1072#1081#1083#1099'|*.*'
    Title = #1060#1072#1081#1083' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080
    Left = 256
    Top = 120
  end
  object SaveDlg: TSaveDialog
    DefaultExt = 'set1'
    Filter = #1060#1072#1081#1083#1099' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080'|*.set1|'#1042#1089#1077' '#1092#1072#1081#1083#1099'|*.*'
    Title = #1060#1072#1081#1083' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080
    Left = 288
    Top = 120
  end
end
