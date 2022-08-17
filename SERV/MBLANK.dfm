object FBLANK: TFBLANK
  Left = 1112
  Top = 223
  BorderStyle = bsDialog
  BorderWidth = 6
  Caption = #1069#1082#1089#1087#1086#1088#1090' '#1073#1083#1072#1085#1082#1086#1074' '#1086#1090#1095#1077#1090#1086#1074
  ClientHeight = 233
  ClientWidth = 473
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnCreate = FormCreate
  OnHide = FormHide
  PixelsPerInch = 96
  TextHeight = 13
  object LCaption: TLabel
    AlignWithMargins = True
    Left = 0
    Top = 3
    Width = 473
    Height = 16
    Margins.Left = 0
    Margins.Right = 0
    Align = alTop
    AutoSize = False
    Caption = #1054#1090#1095#1077#1090#1099' '#1076#1083#1103' '#1101#1082#1089#1087#1086#1088#1090#1072
    ExplicitLeft = -26
    ExplicitWidth = 499
  end
  object CBView: TCheckListBox
    Left = 0
    Top = 22
    Width = 473
    Height = 151
    Align = alClient
    ItemHeight = 13
    Items.Strings = (
      'cvbxcvb'
      'xcvbxcvb'
      'xcvbx'
      'cvbx'
      'xcv')
    TabOrder = 0
    OnClick = CBViewClick
    OnKeyDown = CBViewKeyDown
  end
  object PSpace1: TPanel
    Left = 0
    Top = 173
    Width = 473
    Height = 7
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
    object Panel2: TPanel
      Left = 0
      Top = 0
      Width = 473
      Height = 7
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 0
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 201
    Width = 473
    Height = 7
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 2
  end
  object Panel3: TPanel
    Left = 0
    Top = 208
    Width = 473
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 3
    object BtnCancel: TBitBtn
      Left = 323
      Top = 0
      Width = 150
      Height = 25
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 0
    end
    object Panel4: TPanel
      Left = 304
      Top = 0
      Width = 19
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
    end
    object BtnOk: TBitBtn
      Left = 154
      Top = 0
      Width = 150
      Height = 25
      Action = AOk
      Align = alRight
      Caption = #1047#1072#1087#1080#1089#1072#1090#1100
      TabOrder = 2
    end
  end
  object Panel5: TPanel
    Left = 0
    Top = 180
    Width = 473
    Height = 21
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 4
    object LFolder: TLabel
      AlignWithMargins = True
      Left = 0
      Top = 3
      Width = 57
      Height = 15
      Margins.Left = 0
      Margins.Right = 0
      Align = alLeft
      AutoSize = False
      Caption = #1050#1091#1076#1072
      ExplicitHeight = 18
    end
    object CBFolder: TButtonedEdit
      Left = 57
      Top = 0
      Width = 416
      Height = 21
      Align = alClient
      Images = FMAIN.ImgLst16
      ParentShowHint = False
      RightButton.Hint = #1042#1099#1073#1088#1072#1090#1100' ...'
      RightButton.ImageIndex = 0
      RightButton.Visible = True
      ShowHint = True
      TabOrder = 0
      OnChange = CBFolderChange
      OnRightButtonClick = CBFolderRightButtonClick
    end
  end
  object AList: TActionList
    Images = FMAIN.ImgLstSys
    Left = 360
    Top = 56
    object AOk: TAction
      Caption = #1047#1072#1087#1080#1089#1072#1090#1100
      ImageIndex = 0
      OnExecute = AOkExecute
    end
    object ACancel: TAction
      Caption = #1054#1090#1084#1077#1085#1072
      ImageIndex = 1
      OnExecute = ACancelExecute
    end
  end
end
