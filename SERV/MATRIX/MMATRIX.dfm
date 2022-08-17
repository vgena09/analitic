object FMATRIX: TFMATRIX
  Left = 257
  Top = 162
  ActiveControl = ElPopupButton2
  BorderStyle = bsDialog
  BorderWidth = 6
  Caption = #1052#1072#1089#1090#1077#1088' '#1089#1086#1079#1076#1072#1085#1080#1103' '#1084#1072#1090#1088#1080#1094#1099' '#1092#1086#1088#1084#1099' '#1089#1090#1072#1090#1080#1089#1090#1080#1082#1080
  ClientHeight = 192
  ClientWidth = 618
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
  object PBottom: TElPanel
    Left = 0
    Top = 167
    Width = 618
    Height = 25
    Resizable = True
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 3
    DockOrientation = doNoOrient
    object ElPopupButton1: TElPopupButton
      Left = 416
      Top = 0
      Width = 202
      Height = 25
      ImageIndex = 1
      UseImageList = True
      Images = FMAIN.ImgLstSys
      DrawDefaultFrame = False
      ModalResult = 2
      Caption = #1047#1072#1082#1088#1099#1090#1100
      TabOrder = 0
      Align = alRight
    end
    object ElPanel5: TElPanel
      Left = 0
      Top = 0
      Width = 12
      Height = 25
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 1
      DockOrientation = doNoOrient
    end
    object ElPanel14: TElPanel
      Left = 404
      Top = 0
      Width = 12
      Height = 25
      Align = alRight
      BevelOuter = bvNone
      TabOrder = 2
      DockOrientation = doNoOrient
    end
  end
  object ElPanel1: TElPanel
    Left = 0
    Top = 81
    Width = 618
    Height = 6
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 5
    DockOrientation = doNoOrient
  end
  object ElPanel10: TElPanel
    Left = 0
    Top = 27
    Width = 618
    Height = 6
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 6
    DockOrientation = doNoOrient
  end
  object PBlockKod: TElPanel
    Left = 0
    Top = 87
    Width = 618
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 2
    DockOrientation = doNoOrient
    object LBlockKod: TLMDSimpleLabel
      Left = 0
      Top = 0
      Width = 153
      Height = 21
      Alignment = agCenterLeft
      AutoSize = False
      Align = alLeft
      Caption = #1050#1086#1076' '#1074#1077#1088#1093#1085#1077#1081' '#1083#1077#1074#1086#1081' '#1103#1095#1077#1081#1082#1080
      Options = []
    end
    object EBlockKod: TElButtonEdit
      Left = 153
      Top = 0
      Width = 465
      Height = 21
      UseCustomScrollBars = False
      LineBorderActiveColor = clBlack
      LineBorderInactiveColor = clBlack
      ButtonCaption = #1048#1079#1084#1077#1085#1080#1090#1100' '#1082#1086#1076' '#1103#1095#1077#1081#1082#1080
      ButtonUsePng = False
      ButtonWidth = 160
      OnButtonClick = EBlockKodButtonClick
      AltButtonWidth = 15
      AltButtonUsePng = False
      TabOrder = 0
      Align = alClient
      ParentShowHint = False
      ShowHint = False
    end
  end
  object ElPanel2: TElPanel
    Left = 0
    Top = 0
    Width = 618
    Height = 6
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 7
    DockOrientation = doNoOrient
  end
  object PPage: TElPanel
    Left = 0
    Top = 6
    Width = 618
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 4
    DockOrientation = doNoOrient
    object LPage: TLMDSimpleLabel
      Left = 0
      Top = 0
      Width = 153
      Height = 21
      Alignment = agCenterLeft
      AutoSize = False
      Align = alLeft
      Caption = #1048#1085#1076#1077#1082#1089' '#1083#1080#1089#1090#1072
      Options = []
    end
    object EPage: TElSpinEdit
      Left = 153
      Top = 0
      Width = 60
      Height = 21
      Cursor = crIBeam
      Hint = #1048#1085#1076#1077#1082#1089' ('#1087#1086#1088#1103#1076#1082#1072#1074#1099#1081' '#1085#1086#1084#1077#1088') '#1083#1080#1089#1090#1072
      Value = 1
      MaxValue = 1000
      MinValue = 1
      UseCustomScrollBars = False
      MaxLength = 255
      Align = alLeft
      LineBorderActiveColor = clBlack
      LineBorderInactiveColor = clBlack
      ParentShowHint = False
      TabOrder = 0
      ShowHint = True
    end
    object ElPanel3: TElPanel
      Left = 213
      Top = 0
      Width = 12
      Height = 21
      Align = alLeft
      BevelOuter = bvNone
      TabOrder = 1
      DockOrientation = doNoOrient
    end
  end
  object ElPanel4: TElPanel
    Left = 0
    Top = 108
    Width = 618
    Height = 6
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 8
    DockOrientation = doNoOrient
  end
  object PBlockIn: TElPanel
    Left = 0
    Top = 33
    Width = 618
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    DockOrientation = doNoOrient
    object LBlockIn: TLMDSimpleLabel
      Left = 0
      Top = 0
      Width = 153
      Height = 21
      Alignment = agCenterLeft
      AutoSize = False
      Align = alLeft
      Caption = #1041#1083#1086#1082' '#1074#1074#1086#1076#1072
      Options = []
    end
    object EBlockIn: TElButtonEdit
      Left = 153
      Top = 0
      Width = 465
      Height = 21
      Hint = 
        #1040#1076#1088#1077#1089#1072' '#1103#1095#1077#1077#1082', '#1080#1079' '#1082#1086#1090#1086#1088#1099#1093' '#1086#1089#1091#1097#1077#1089#1090#1074#1083#1103#1077#1090#1089#1103' '#1074#1074#1086#1076' '#1080#1085#1092#1086#1088#1084#1072#1094#1080#1080' '#1074' '#1073#1072#1079#1091' '#1076 +
        #1072#1085#1085#1099#1093
      UseCustomScrollBars = False
      LineBorderActiveColor = clBlack
      LineBorderInactiveColor = clBlack
      ButtonCaption = #1057#1080#1085#1093#1088#1086#1085
      ButtonHint = #1057#1080#1085#1093#1088#1086#1085#1080#1079#1080#1088#1086#1074#1072#1090#1100' '#1080#1079#1084#1077#1085#1077#1085#1080#1103' '#1073#1083#1086#1082#1072' '#1074#1074#1086#1076#1072' '#1089' '#1073#1083#1086#1082#1086#1084' '#1074#1099#1074#1086#1076#1072
      ButtonIsSwitch = True
      ButtonUsePng = False
      ButtonWidth = 80
      OnButtonClick = EBlockInButtonClick
      AltButtonCaption = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100
      AltButtonHint = #1050#1086#1087#1080#1088#1086#1074#1072#1090#1100' '#1073#1083#1086#1082' '#1074#1074#1086#1076#1072' '#1074' '#1073#1083#1086#1082' '#1074#1099#1074#1086#1076#1072
      AltButtonVisible = True
      AltButtonWidth = 80
      AltButtonUsePng = False
      OnAltButtonClick = EBlockInAltButtonClick
      TabOrder = 0
      Align = alClient
      ParentShowHint = False
      ShowHint = True
      OnEnter = EBlockInEnter
      OnExit = EBlockInExit
    end
  end
  object ElPanel11: TElPanel
    Left = 0
    Top = 54
    Width = 618
    Height = 6
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 9
    DockOrientation = doNoOrient
  end
  object PBlockOut: TElPanel
    Left = 0
    Top = 60
    Width = 618
    Height = 21
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 1
    DockOrientation = doNoOrient
    object LBlockOut: TLMDSimpleLabel
      Left = 0
      Top = 0
      Width = 153
      Height = 21
      Alignment = agCenterLeft
      AutoSize = False
      Align = alLeft
      Caption = #1041#1083#1086#1082#1080' '#1074#1099#1074#1086#1076#1072
      Options = []
    end
    object EBlockOut: TElButtonEdit
      Left = 153
      Top = 0
      Width = 465
      Height = 21
      Hint = #1040#1076#1088#1077#1089#1072' '#1103#1095#1077#1077#1082', '#1074' '#1082#1086#1090#1086#1088#1099#1077' '#1088#1072#1079#1084#1077#1097#1072#1077#1090#1089#1103' '#1080#1085#1092#1086#1088#1084#1072#1094#1080#1103' '#1080#1079' '#1073#1072#1079#1099' '#1076#1072#1085#1085#1099#1093
      UseCustomScrollBars = False
      LineBorderActiveColor = clBlack
      LineBorderInactiveColor = clBlack
      ButtonCaption = #1054#1095#1080#1089#1090#1080#1090#1100
      ButtonUsePng = False
      ButtonWidth = 80
      OnButtonClick = EBlockOutButtonClick
      AltButtonCaption = #1057#1083#1077#1076#1091#1102#1097#1080#1081
      AltButtonVisible = True
      AltButtonWidth = 80
      AltButtonUsePng = False
      OnAltButtonClick = EBlockOutAltButtonClick
      TabOrder = 0
      Align = alClient
      ParentShowHint = False
      ShowHint = True
      OnEnter = EBlockOutEnter
      OnExit = EBlockOutExit
    end
  end
  object ElPopupButton2: TElPopupButton
    Left = 0
    Top = 128
    Width = 200
    Height = 25
    ImageIndex = 0
    UseImageList = True
    Images = FMAIN.ImgLstSys
    DrawDefaultFrame = False
    TabOrder = 10
    Action = ARun
  end
  object ElPopupButton3: TElPopupButton
    Left = 208
    Top = 128
    Width = 201
    Height = 25
    ImageIndex = 4
    UseImageList = True
    Images = FMAIN.ImgLstSys
    DrawDefaultFrame = False
    TabOrder = 11
    Action = AClear
  end
  object ElPopupButton4: TElPopupButton
    Left = 416
    Top = 128
    Width = 200
    Height = 25
    ImageIndex = 2
    UseImageList = True
    Images = FMAIN.ImgLstSys
    DrawDefaultFrame = False
    TabOrder = 12
    Action = AOpen
  end
  object XLS: TExcelApplication
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    AutoQuit = False
    Left = 432
  end
  object WB: TExcelWorkbook
    AutoConnect = False
    ConnectKind = ckRunningOrNew
    OnBeforeClose = WBBeforeClose
    OnSheetSelectionChange = WBSheetSelectionChange
    Left = 464
  end
  object ActionList1: TActionList
    Images = FMAIN.ImgLstSys
    Left = 400
    object ARun: TAction
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1073#1083#1086#1082
      ImageIndex = 0
      OnExecute = ARunExecute
    end
    object AClear: TAction
      Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1089#1077#1082#1094#1080#1102
      Hint = #1054#1095#1080#1089#1090#1080#1090#1100' '#1084#1072#1090#1088#1080#1094#1091' '#1083#1080#1089#1090#1072
      ImageIndex = 4
      OnExecute = AClearExecute
    end
    object AOpen: TAction
      Caption = #1054#1090#1082#1088#1099#1090#1100' '#1084#1072#1090#1088#1080#1094#1091
      ImageIndex = 2
      OnExecute = AOpenExecute
    end
  end
end
