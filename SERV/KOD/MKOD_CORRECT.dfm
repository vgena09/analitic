object FKOD_CORRECT: TFKOD_CORRECT
  Left = 799
  Top = 330
  BorderStyle = bsDialog
  Caption = #1050#1086#1088#1088#1077#1082#1094#1080#1103' '#1082#1086#1076#1072
  ClientHeight = 154
  ClientWidth = 750
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
  object Bevel2: TBevel
    Left = 0
    Top = 83
    Width = 750
    Height = 31
    Align = alBottom
    Shape = bsBottomLine
    ExplicitTop = 134
    ExplicitWidth = 707
  end
  object Bevel3: TBevel
    Left = 0
    Top = 0
    Width = 750
    Height = 8
    Align = alTop
    Shape = bsSpacer
    ExplicitWidth = 373
  end
  object POld: TPanel
    Left = 0
    Top = 8
    Width = 750
    Height = 33
    Align = alTop
    BevelOuter = bvNone
    BorderWidth = 6
    TabOrder = 0
    object Label1: TLabel
      AlignWithMargins = True
      Left = 9
      Top = 9
      Width = 136
      Height = 15
      Align = alLeft
      AutoSize = False
      Caption = #1060#1086#1088#1084#1091#1083#1072' '#1080#1083#1080' '#1082#1086#1076
      ExplicitHeight = 14
    end
    object BtnOld: TButton
      Left = 591
      Top = 6
      Width = 70
      Height = 21
      Action = AKodFormula
      Align = alRight
      TabOrder = 0
    end
    object Panel1: TPanel
      Left = 578
      Top = 6
      Width = 13
      Height = 21
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 1
    end
    object Button1: TButton
      Left = 674
      Top = 6
      Width = 70
      Height = 21
      Action = AKodOld
      Align = alRight
      TabOrder = 2
    end
    object Panel3: TPanel
      Left = 661
      Top = 6
      Width = 13
      Height = 21
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 3
    end
    object CBFormula: TComboBox
      Left = 148
      Top = 6
      Width = 430
      Height = 21
      Align = alClient
      TabOrder = 4
      Text = 'CBFormula'
      OnChange = EditChange
    end
  end
  object PNew: TPanel
    Left = 0
    Top = 51
    Width = 750
    Height = 32
    Align = alBottom
    BevelOuter = bvNone
    BorderWidth = 6
    TabOrder = 1
    object Label4: TLabel
      AlignWithMargins = True
      Left = 9
      Top = 9
      Width = 136
      Height = 14
      Align = alLeft
      AutoSize = False
      Caption = #1047#1072#1087#1080#1089#1072#1090#1100' '#1088#1077#1079#1091#1083#1100#1090#1072#1090' '#1074
    end
    object Label5: TLabel
      AlignWithMargins = True
      Left = 196
      Top = 9
      Width = 17
      Height = 13
      Align = alLeft
      Caption = '  /  '
    end
    object Label6: TLabel
      AlignWithMargins = True
      Left = 264
      Top = 9
      Width = 17
      Height = 13
      Align = alLeft
      Caption = '  /  '
    end
    object ETab2: TEdit
      Left = 148
      Top = 6
      Width = 45
      Height = 20
      Align = alLeft
      TabOrder = 0
      OnChange = EditChange
      ExplicitHeight = 21
    end
    object ECol2: TEdit
      Left = 216
      Top = 6
      Width = 45
      Height = 20
      Align = alLeft
      TabOrder = 1
      OnChange = EditChange
      ExplicitHeight = 21
    end
    object ERow2: TEdit
      Left = 284
      Top = 6
      Width = 45
      Height = 20
      Align = alLeft
      TabOrder = 2
      OnChange = EditChange
      ExplicitHeight = 21
    end
    object BtnNewKod: TButton
      Left = 674
      Top = 6
      Width = 70
      Height = 20
      Action = AKodNew
      Align = alRight
      TabOrder = 3
    end
  end
  object PBottom: TPanel
    Left = 0
    Top = 114
    Width = 750
    Height = 40
    Align = alBottom
    BevelOuter = bvNone
    BorderWidth = 6
    TabOrder = 2
    object BitBtn1: TBitBtn
      Left = 325
      Top = 6
      Width = 200
      Height = 28
      Action = ARun
      Align = alRight
      Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100
      TabOrder = 0
    end
    object BitBtn2: TBitBtn
      Left = 544
      Top = 6
      Width = 200
      Height = 28
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 1
    end
    object Panel2: TPanel
      Left = 525
      Top = 6
      Width = 19
      Height = 28
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 2
    end
  end
  object ActionList1: TActionList
    Images = FMAIN.ImgLstSys
    Left = 184
    Top = 104
    object ARun: TAction
      Caption = #1042#1099#1087#1086#1083#1085#1080#1090#1100
      ImageIndex = 0
      OnExecute = ARunExecute
    end
    object ACancel: TAction
      Caption = #1054#1090#1084#1077#1085#1072
      ImageIndex = 1
      OnExecute = ACancelExecute
    end
    object AKodFormula: TAction
      Tag = 1
      Caption = #1060#1086#1088#1084#1091#1083#1072
      OnExecute = AKodFormulaExecute
    end
    object AKodOld: TAction
      Caption = #1050#1086#1076
      OnExecute = AKodOldExecute
    end
    object AKodNew: TAction
      Tag = 2
      Caption = #1050#1086#1076
      OnExecute = AKodNewExecute
    end
  end
end
