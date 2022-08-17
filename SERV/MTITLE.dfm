object FTITLE: TFTITLE
  Left = 361
  Top = 164
  BorderStyle = bsDialog
  Caption = #1057#1074#1077#1076#1077#1085#1080#1103' '#1086#1073' '#1086#1090#1095#1077#1090#1077
  ClientHeight = 170
  ClientWidth = 500
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
  OnHide = FormHide
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    AlignWithMargins = True
    Left = 3
    Top = 134
    Width = 494
    Height = 2
    Align = alBottom
    ExplicitLeft = 0
    ExplicitTop = 114
    ExplicitWidth = 488
  end
  object Panel1: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 142
    Width = 494
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object SpeedButton1: TSpeedButton
      Left = 146
      Top = -41
      Width = 23
      Height = 22
    end
    object BtnCancel: TBitBtn
      Left = 344
      Top = 0
      Width = 150
      Height = 25
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 2
    end
    object BtnOk: TBitBtn
      Left = 0
      Top = 0
      Width = 150
      Height = 25
      Action = AOK
      Align = alLeft
      Caption = #1054#1082
      TabOrder = 0
    end
    object Panel2: TPanel
      Left = 150
      Top = 0
      Width = 13
      Height = 25
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 3
    end
    object Panel3: TPanel
      Left = 331
      Top = 0
      Width = 13
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 4
    end
    object BtnRestore: TBitBtn
      Left = 163
      Top = 0
      Width = 168
      Height = 25
      Action = ARestore
      Align = alClient
      Caption = #1042#1086#1089#1089#1090#1072#1085#1086#1074#1080#1090#1100' '#1079#1085#1072#1095#1077#1085#1080#1103
      TabOrder = 1
    end
  end
  object PPeriod: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 494
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object LPeriod: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 97
      Height = 15
      Align = alLeft
      AutoSize = False
      Caption = #1055#1077#1088#1080#1086#1076
    end
    object CBYear: TComboBox
      Left = 103
      Top = 0
      Width = 70
      Height = 21
      Align = alLeft
      DropDownCount = 10
      TabOrder = 0
      OnChange = CBPeriodChange
      OnKeyDown = FormKeyDown
    end
    object LPeriodSpace: TPanel
      Left = 173
      Top = 0
      Width = 13
      Height = 21
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 1
    end
    object CBMonth: TComboBox
      Left = 186
      Top = 0
      Width = 130
      Height = 21
      Align = alLeft
      DropDownCount = 12
      TabOrder = 2
      OnChange = CBPeriodChange
      OnKeyDown = FormKeyDown
    end
  end
  object PForm: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 30
    Width = 494
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 2
    object LForm: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 97
      Height = 15
      Align = alLeft
      AutoSize = False
      Caption = #1054#1090#1095#1077#1090
      ExplicitLeft = 0
      ExplicitTop = 6
    end
    object CBForm: TComboBox
      Left = 103
      Top = 0
      Width = 391
      Height = 21
      Align = alClient
      DropDownCount = 30
      TabOrder = 0
      OnChange = CBFormChange
      OnKeyDown = FormKeyDown
    end
  end
  object PRegion: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 57
    Width = 494
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 3
    object LReg: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 97
      Height = 15
      Align = alLeft
      AutoSize = False
      Caption = #1056#1077#1075#1080#1086#1085
      ExplicitLeft = 0
      ExplicitTop = 6
    end
    object CBRegion: TComboBox
      Left = 103
      Top = 0
      Width = 391
      Height = 21
      Align = alClient
      DropDownCount = 30
      TabOrder = 0
      OnChange = CBRegionChange
      OnKeyDown = FormKeyDown
    end
  end
  object PSubRegion: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 84
    Width = 494
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 4
    object LSubRegion: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 97
      Height = 15
      Align = alLeft
      AutoSize = False
      Caption = #1055#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1077
      ExplicitLeft = 0
      ExplicitTop = 6
    end
    object CBSubRegion: TEdit
      Left = 103
      Top = 0
      Width = 391
      Height = 21
      Align = alClient
      TabOrder = 0
      OnChange = CBSubRegionChange
      OnKeyDown = FormKeyDown
    end
  end
  object CBPrev: TCheckBox
    AlignWithMargins = True
    Left = 3
    Top = 111
    Width = 494
    Height = 17
    Align = alTop
    Caption = #1048#1084#1087#1086#1088#1090#1080#1088#1086#1074#1072#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1087#1088#1077#1076#1099#1076#1091#1097#1077#1075#1086' '#1087#1077#1088#1080#1086#1076#1072
    TabOrder = 5
  end
  object ActionList1: TActionList
    Images = FMAIN.ImgLstSys
    Left = 376
    object AOK: TAction
      Caption = #1054#1082
      ImageIndex = 0
      OnExecute = AOKExecute
    end
    object ACancel: TAction
      Caption = #1054#1090#1084#1077#1085#1072
      ImageIndex = 1
      OnExecute = ACancelExecute
    end
    object ARestore: TAction
      Caption = #1042#1086#1089#1089#1090#1072#1085#1086#1074#1080#1090#1100' '#1079#1085#1072#1095#1077#1085#1080#1103
      ImageIndex = 2
      OnExecute = ARestoreExecute
    end
    object ASave: TAction
      Caption = #1057#1086#1093#1088#1072#1085#1077#1085#1080#1077' '#1072#1074#1090#1086#1079#1072#1084#1077#1085#1099' '#1088#1077#1075#1080#1086#1085#1072
      Hint = #1057#1086#1093#1088#1072#1085#1077#1085#1080#1077' '#1072#1074#1090#1086#1079#1072#1084#1077#1085#1099' '#1088#1077#1075#1080#1086#1085#1072
      ImageIndex = 3
      ShortCut = 16467
      OnExecute = ASaveExecute
    end
  end
end
