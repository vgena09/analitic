object FBDREGION: TFBDREGION
  Left = 432
  Top = 249
  BorderStyle = bsDialog
  Caption = #1057#1086#1079#1076#1072#1090#1100' '#1041#1044' '#1088#1077#1075#1080#1086#1085#1072
  ClientHeight = 90
  ClientWidth = 365
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
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object PRegion: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 3
    Width = 359
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    ExplicitLeft = 0
    ExplicitTop = 48
    ExplicitWidth = 353
    object LRegion: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 67
      Height = 15
      Align = alLeft
      AutoSize = False
      Caption = #1056#1077#1075#1080#1086#1085
    end
    object CBRegion: TComboBox
      Left = 73
      Top = 0
      Width = 286
      Height = 21
      Align = alClient
      DropDownCount = 30
      TabOrder = 0
      OnChange = CBRegionChange
    end
  end
  object PFolder: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 30
    Width = 359
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    ExplicitLeft = 0
    ExplicitTop = 56
    ExplicitWidth = 365
    object LFolder: TLabel
      AlignWithMargins = True
      Left = 3
      Top = 3
      Width = 67
      Height = 15
      Align = alLeft
      AutoSize = False
      Caption = #1050#1091#1076#1072
    end
    object CBFolder: TButtonedEdit
      Left = 73
      Top = 0
      Width = 286
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
      ExplicitLeft = 70
      ExplicitWidth = 283
    end
  end
  object Panel3: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 62
    Width = 359
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 2
    ExplicitLeft = -108
    ExplicitTop = 150
    ExplicitWidth = 473
    object BitCancel: TBitBtn
      Left = 189
      Top = 0
      Width = 170
      Height = 25
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 0
      ExplicitLeft = 215
    end
    object BtnOk: TBitBtn
      Left = 0
      Top = 0
      Width = 170
      Height = 25
      Action = AOK
      Align = alLeft
      Caption = #1054#1082
      TabOrder = 1
      ExplicitTop = 3
    end
  end
  object ActionList1: TActionList
    Images = FMAIN.ImgLstSys
    Left = 176
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
  end
end
