object FBANK: TFBANK
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = #1055#1088#1080#1086#1088#1080#1090#1077#1090' '#1073#1072#1085#1082#1086#1074' '#1076#1072#1085#1085#1099#1093
  ClientHeight = 164
  ClientWidth = 253
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object PBottom: TPanel
    AlignWithMargins = True
    Left = 3
    Top = 136
    Width = 247
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    ExplicitTop = 179
    object BitBtn1: TBitBtn
      Left = 0
      Top = 0
      Width = 120
      Height = 25
      Action = AOk
      Align = alLeft
      Caption = #1054#1082
      TabOrder = 0
    end
    object BitBtn2: TBitBtn
      Left = 127
      Top = 0
      Width = 120
      Height = 25
      Action = ACancel
      Align = alRight
      Caption = #1054#1090#1084#1077#1085#1072
      TabOrder = 1
    end
  end
  object LBox: TListBox
    Left = 0
    Top = 0
    Width = 253
    Height = 133
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 1
    ExplicitHeight = 176
  end
  object ActionList1: TActionList
    Images = FMAIN.ImgLstSys
    Left = 8
    Top = 8
    object AOk: TAction
      Caption = #1054#1082
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
