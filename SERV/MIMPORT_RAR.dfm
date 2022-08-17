object FIMPORT_RAR: TFIMPORT_RAR
  Left = 616
  Top = 265
  BorderWidth = 6
  Caption = #1048#1084#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093
  ClientHeight = 599
  ClientWidth = 499
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object LComment: TLabel
    AlignWithMargins = True
    Left = 0
    Top = 410
    Width = 499
    Height = 20
    Margins.Left = 0
    Margins.Right = 0
    Align = alBottom
    AutoSize = False
    Caption = #1050#1086#1084#1084#1077#1085#1090#1072#1088#1080#1081' '#1086#1090#1087#1088#1072#1074#1080#1090#1077#1083#1103
    ExplicitTop = 332
  end
  object LCaption: TLabel
    AlignWithMargins = True
    Left = 0
    Top = 28
    Width = 499
    Height = 16
    Margins.Left = 0
    Margins.Right = 0
    Align = alTop
    AutoSize = False
    Caption = #1042#1099#1073#1088#1072#1085#1086' 0 '#1088#1077#1075'. '#1080#1079' 0'
    ExplicitTop = 3
  end
  object CBView: TCheckListBox
    Left = 0
    Top = 47
    Width = 499
    Height = 353
    Align = alClient
    ItemHeight = 13
    Items.Strings = (
      'cvbxcvb'
      'xcvbxcvb'
      'xcvbx'
      'cvbx'
      'xcv')
    PopupMenu = PMenu
    TabOrder = 0
    OnClick = CBViewClick
    OnKeyDown = CBViewKeyDown
  end
  object CBComment: TMemo
    Left = 0
    Top = 433
    Width = 499
    Height = 86
    Align = alBottom
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlue
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    Lines.Strings = (
      'CBComment')
    ParentFont = False
    TabOrder = 1
  end
  object CBSubReg: TCheckBox
    Left = 0
    Top = 526
    Width = 499
    Height = 17
    Align = alBottom
    Caption = #1048#1084#1087#1086#1088#1090#1080#1088#1086#1074#1072#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1089#1090#1088#1091#1082#1090#1091#1088#1085#1099#1093' '#1087#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1081
    TabOrder = 2
    OnClick = CBSubRegClick
    OnKeyPress = CBSubRegKeyPress
  end
  object CBVerify: TCheckBox
    Left = 0
    Top = 550
    Width = 499
    Height = 17
    Align = alBottom
    Caption = #1055#1088#1086#1074#1077#1088#1080#1090#1100' '#1080#1084#1087#1086#1088#1090#1080#1088#1086#1074#1072#1085#1085#1099#1077' '#1086#1090#1095#1077#1090#1099
    TabOrder = 3
    OnClick = CBVerifyClick
    OnKeyPress = CBVerifyKeyPress
  end
  object PBottom: TPanel
    Left = 0
    Top = 574
    Width = 499
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 4
    object Panel2: TPanel
      Left = 339
      Top = 0
      Width = 20
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
    end
    object BtnOk: TBitBtn
      Left = 199
      Top = 0
      Width = 140
      Height = 25
      Action = AImport
      Align = alRight
      Caption = #1048#1084#1087#1086#1088#1090
      TabOrder = 1
    end
    object BtnCancel: TBitBtn
      Left = 359
      Top = 0
      Width = 140
      Height = 25
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 2
    end
  end
  object PSpaceComment: TPanel
    Left = 0
    Top = 400
    Width = 499
    Height = 7
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 5
  end
  object Panel1: TPanel
    Left = 0
    Top = 567
    Width = 499
    Height = 7
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 6
  end
  object Panel3: TPanel
    Left = 0
    Top = 519
    Width = 499
    Height = 7
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 7
  end
  object Panel4: TPanel
    Left = 0
    Top = 543
    Width = 499
    Height = 7
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 8
  end
  object PVersion: TPanel
    Left = 0
    Top = 0
    Width = 499
    Height = 25
    Align = alTop
    BevelOuter = bvLowered
    Caption = 'PVersion'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clRed
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 9
  end
  object ActionList: TActionList
    Images = FMAIN.ImgLstSys
    Left = 344
    Top = 40
    object AImport: TAction
      Caption = #1048#1084#1087#1086#1088#1090
      ImageIndex = 0
      OnExecute = AImportExecute
    end
    object ACancel: TAction
      Caption = #1054#1090#1084#1077#1085#1072
      ImageIndex = 1
      OnExecute = ACancelExecute
    end
    object ASelectAll: TAction
      Caption = #1042#1099#1076#1077#1083#1080#1090#1100' '#1074#1089#1077
      OnExecute = ASelectAllExecute
    end
    object AUnselectAll: TAction
      Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1074#1089#1077
      OnExecute = AUnselectAllExecute
    end
  end
  object PMenu: TPopupMenu
    Left = 376
    Top = 40
    object N1: TMenuItem
      Action = ASelectAll
    end
    object N2: TMenuItem
      Action = AUnselectAll
    end
  end
end
