object FEXPORT: TFEXPORT
  Left = 679
  Top = 439
  BorderWidth = 6
  Caption = #1069#1082#1089#1087#1086#1088#1090' '#1076#1072#1085#1085#1099#1093
  ClientHeight = 330
  ClientWidth = 572
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
  object PPeriod: TPanel
    Left = 0
    Top = 0
    Width = 572
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object LPeriod: TLabel
      AlignWithMargins = True
      Left = 0
      Top = 3
      Width = 80
      Height = 15
      Margins.Left = 0
      Margins.Right = 0
      Align = alLeft
      AutoSize = False
      Caption = #1055#1077#1088#1080#1086#1076
      ExplicitTop = 6
    end
    object Label1: TLabel
      AlignWithMargins = True
      Left = 185
      Top = 3
      Width = 10
      Height = 15
      Margins.Left = 0
      Margins.Right = 0
      Align = alLeft
      AutoSize = False
      ExplicitLeft = 155
      ExplicitHeight = 18
    end
    object CBYear: TComboBox
      Left = 80
      Top = 0
      Width = 105
      Height = 21
      Align = alLeft
      TabOrder = 0
      Text = 'CBYear'
      OnChange = CBYearChange
    end
    object CBMonth: TComboBox
      Left = 195
      Top = 0
      Width = 377
      Height = 21
      Align = alClient
      DropDownCount = 13
      TabOrder = 1
      Text = 'CBMonth'
      OnChange = CBMonthChange
    end
  end
  object PFolder: TPanel
    Left = 0
    Top = 28
    Width = 572
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    object LFolder: TLabel
      AlignWithMargins = True
      Left = 0
      Top = 3
      Width = 80
      Height = 15
      Margins.Left = 0
      Margins.Right = 0
      Align = alLeft
      AutoSize = False
      Caption = #1050#1091#1076#1072
      ExplicitTop = 6
    end
    object CBFolder: TButtonedEdit
      Left = 80
      Top = 0
      Width = 492
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
  object PComment: TPanel
    Left = 0
    Top = 56
    Width = 572
    Height = 242
    Align = alClient
    BevelOuter = bvNone
    TabOrder = 2
    object LComment: TLabel
      AlignWithMargins = True
      Left = 0
      Top = 3
      Width = 80
      Height = 236
      Margins.Left = 0
      Margins.Right = 0
      Align = alLeft
      AutoSize = False
      Caption = #1050#1086#1084#1084#1077#1085#1090#1072#1088#1080#1081
      ExplicitLeft = -3
      ExplicitTop = 0
      ExplicitHeight = 99
    end
    object CBComment: TMemo
      Left = 80
      Top = 0
      Width = 492
      Height = 242
      Align = alClient
      TabOrder = 0
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 21
    Width = 572
    Height = 7
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 3
  end
  object Panel2: TPanel
    Left = 0
    Top = 49
    Width = 572
    Height = 7
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 4
  end
  object Panel3: TPanel
    Left = 0
    Top = 298
    Width = 572
    Height = 7
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 5
  end
  object PBottom: TPanel
    Left = 0
    Top = 305
    Width = 572
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 6
    object Panel4: TPanel
      Left = 386
      Top = 0
      Width = 26
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 0
    end
    object BtnCancel: TBitBtn
      Left = 412
      Top = 0
      Width = 160
      Height = 25
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 1
    end
    object BtnOk: TBitBtn
      Left = 226
      Top = 0
      Width = 160
      Height = 25
      Action = AExport
      Align = alRight
      Caption = #1069#1082#1089#1087#1086#1088#1090
      TabOrder = 2
    end
  end
  object ActionList: TActionList
    Images = FMAIN.ImgLstSys
    Left = 8
    Top = 88
    object AExport: TAction
      Caption = #1069#1082#1089#1087#1086#1088#1090
      ImageIndex = 0
      OnExecute = AExportExecute
    end
    object ACancel: TAction
      Caption = #1054#1090#1084#1077#1085#1072
      ImageIndex = 1
      OnExecute = ACancelExecute
    end
  end
end
