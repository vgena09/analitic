object FSET: TFSET
  Left = 642
  Top = 195
  BorderWidth = 6
  Caption = #1053#1072#1089#1090#1088#1086#1081#1082#1080
  ClientHeight = 283
  ClientWidth = 472
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
  object PBottom: TPanel
    Left = 0
    Top = 258
    Width = 472
    Height = 25
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 0
    object BtnClose: TBitBtn
      Left = 242
      Top = 0
      Width = 230
      Height = 25
      Align = alRight
      Caption = #1047#1072#1082#1088#1099#1090#1100
      ModalResult = 1
      TabOrder = 0
    end
    object BtnReset: TBitBtn
      Left = 0
      Top = 0
      Width = 230
      Height = 25
      Align = alLeft
      Caption = #1057#1073#1088#1086#1089#1080#1090#1100' '#1085#1072#1089#1090#1088#1086#1081#1082#1080
      TabOrder = 1
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 252
    Width = 472
    Height = 6
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 1
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 472
    Height = 247
    ActivePage = TabSheet2
    Align = alTop
    TabOrder = 2
    object TabSheet1: TTabSheet
      BorderWidth = 6
      Caption = #1054#1089#1085#1086#1074#1085#1099#1077
      object Bevel2: TBevel
        AlignWithMargins = True
        Left = 3
        Top = 76
        Width = 446
        Height = 10
        Align = alTop
        Shape = bsBottomLine
        ExplicitLeft = -3
        ExplicitTop = 115
        ExplicitWidth = 452
      end
      object PMainReg: TPanel
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 446
        Height = 21
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 0
        object LMainReg: TLabel
          AlignWithMargins = True
          Left = 3
          Top = 3
          Width = 100
          Height = 15
          Align = alLeft
          AutoSize = False
          Caption = #1043#1083#1072#1074#1085#1099#1081' '#1088#1077#1075#1080#1086#1085
        end
        object EMainReg: TComboBox
          Left = 106
          Top = 0
          Width = 340
          Height = 21
          Align = alClient
          Style = csDropDownList
          DropDownCount = 30
          TabOrder = 0
          OnChange = EMainRegChange
        end
      end
      object CBNet: TCheckBox
        AlignWithMargins = True
        Left = 3
        Top = 30
        Width = 446
        Height = 17
        Align = alTop
        Caption = #1056#1072#1073#1086#1090#1072' '#1074' '#1089#1077#1090#1080
        TabOrder = 1
        OnClick = CBNetClick
        OnKeyDown = CBNetKeyDown
      end
      object CBTerminate: TCheckBox
        AlignWithMargins = True
        Left = 3
        Top = 53
        Width = 446
        Height = 17
        Align = alTop
        Caption = #1054#1090#1082#1083#1102#1095#1080#1090#1100' '#1087#1086#1083#1100#1079#1086#1074#1072#1090#1077#1083#1077#1081
        TabOrder = 2
        OnClick = CBTerminateClick
        OnKeyDown = CBTerminateKeyDown
      end
      object GroupBox1: TGroupBox
        AlignWithMargins = True
        Left = 3
        Top = 99
        Width = 446
        Height = 105
        Align = alBottom
        Caption = ' '#1056#1072#1073#1086#1090#1072' '#1089' '#1089#1086#1094#1080#1072#1083#1100#1085#1086'-'#1101#1082#1086#1085#1086#1084#1080#1095#1077#1089#1082#1080#1084#1080' '#1087#1086#1082#1072#1079#1072#1090#1077#1083#1103#1084#1080' '
        TabOrder = 3
        object CBSEPDetail: TCheckBox
          AlignWithMargins = True
          Left = 5
          Top = 18
          Width = 436
          Height = 17
          Align = alTop
          Caption = #1054#1087#1090#1080#1084#1080#1079#1080#1088#1086#1074#1072#1090#1100' '#1074#1099#1073#1086#1088' ('#1079#1072#1084#1077#1076#1083#1077#1085#1080#1077' '#1086#1073#1088#1072#1073#1086#1090#1082#1080')'
          TabOrder = 0
          OnClick = CBSEPDetailClick
          OnKeyPress = CBSEPDetailKeyPress
        end
        object Panel3: TPanel
          AlignWithMargins = True
          Left = 5
          Top = 41
          Width = 436
          Height = 21
          Align = alTop
          BevelOuter = bvNone
          TabOrder = 1
          object LSEPMonth: TLabel
            AlignWithMargins = True
            Left = 3
            Top = 3
            Width = 279
            Height = 15
            Align = alLeft
            AutoSize = False
            Caption = #1052#1072#1082#1089#1080#1084#1072#1083#1100#1085#1086#1077' '#1095#1080#1089#1083#1086' '#1087#1088#1086#1089#1084#1072#1090#1088#1080#1074#1072#1077#1084#1099#1093' '#1084#1077#1089#1103#1094#1077#1074
          end
          object LSEPMonth2: TLabel
            AlignWithMargins = True
            Left = 288
            Top = 3
            Width = 21
            Height = 15
            Align = alLeft
            AutoSize = False
            Caption = '11'
          end
          object TBSEPMonth: TTrackBar
            Left = 312
            Top = 0
            Width = 124
            Height = 21
            Align = alClient
            Max = 60
            Min = 12
            Position = 12
            ShowSelRange = False
            TabOrder = 0
            TickStyle = tsNone
            OnChange = TBSEPMonthChange
          end
        end
      end
    end
    object TabSheet2: TTabSheet
      BorderWidth = 6
      Caption = #1069#1082#1089#1087#1086#1088#1090'/'#1048#1084#1087#1086#1088#1090
      ImageIndex = 1
      object Bevel3: TBevel
        AlignWithMargins = True
        Left = 3
        Top = 72
        Width = 446
        Height = 10
        Align = alTop
        Shape = bsTopLine
        ExplicitLeft = -3
        ExplicitTop = 82
        ExplicitWidth = 452
      end
      object Bevel4: TBevel
        AlignWithMargins = True
        Left = 3
        Top = 142
        Width = 446
        Height = 10
        Align = alTop
        Shape = bsTopLine
        ExplicitTop = 168
      end
      object CBImportSubReg: TCheckBox
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 446
        Height = 17
        Align = alTop
        Caption = #1048#1084#1087#1086#1088#1090#1080#1088#1086#1074#1072#1090#1100' '#1076#1072#1085#1085#1099#1077' '#1089#1090#1088#1091#1082#1090#1091#1088#1085#1099#1093' '#1087#1086#1076#1088#1072#1079#1076#1077#1083#1077#1085#1080#1081
        TabOrder = 0
        OnClick = CBImportSubRegClick
        OnKeyPress = CBImportSubRegKeyPress
      end
      object PPathArj: TPanel
        AlignWithMargins = True
        Left = 3
        Top = 88
        Width = 446
        Height = 21
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 1
        object LPathArj: TLabel
          AlignWithMargins = True
          Left = 3
          Top = 3
          Width = 100
          Height = 15
          Align = alLeft
          AutoSize = False
          Caption = #1055#1091#1090#1100' '#1072#1088#1093#1080#1074#1072#1094#1080#1080
        end
        object EPathArj: TButtonedEdit
          Left = 106
          Top = 0
          Width = 340
          Height = 21
          Align = alClient
          Images = FMAIN.ImgLst16
          ParentShowHint = False
          RightButton.Hint = #1042#1099#1073#1088#1072#1090#1100' ...'
          RightButton.ImageIndex = 0
          RightButton.Visible = True
          ShowHint = True
          TabOrder = 0
          OnChange = EPathArjChange
          OnRightButtonClick = EPathArjRightButtonClick
        end
      end
      object CBVerifyRAR: TCheckBox
        AlignWithMargins = True
        Left = 3
        Top = 26
        Width = 446
        Height = 17
        Align = alTop
        Caption = #1055#1088#1086#1074#1077#1088#1103#1090#1100' '#1080#1084#1087#1086#1088#1090#1080#1088#1086#1074#1072#1085#1085#1099#1077' RAR-'#1086#1090#1095#1077#1090#1099
        TabOrder = 2
        OnClick = CBVerifyRARClick
        OnKeyPress = CBVerifyRARKeyPress
      end
      object CBVerifyXLS: TCheckBox
        AlignWithMargins = True
        Left = 3
        Top = 49
        Width = 446
        Height = 17
        Align = alTop
        Caption = #1055#1088#1086#1074#1077#1088#1103#1090#1100' '#1080#1084#1087#1086#1088#1090#1080#1088#1086#1074#1072#1085#1085#1099#1077' XLS-'#1086#1090#1095#1077#1090#1099
        TabOrder = 3
        OnClick = CBVerifyXLSClick
        OnKeyPress = CBVerifyXLSKeyPress
      end
      object PPathExport: TPanel
        AlignWithMargins = True
        Left = 3
        Top = 115
        Width = 446
        Height = 21
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 4
        object LPathExport: TLabel
          AlignWithMargins = True
          Left = 3
          Top = 3
          Width = 100
          Height = 15
          Align = alLeft
          AutoSize = False
          Caption = #1055#1091#1090#1100' '#1101#1082#1089#1087#1086#1088#1090#1072
        end
        object EPathExport: TButtonedEdit
          Left = 106
          Top = 0
          Width = 340
          Height = 21
          Align = alClient
          Images = FMAIN.ImgLst16
          ParentShowHint = False
          RightButton.Hint = #1042#1099#1073#1088#1072#1090#1100' ...'
          RightButton.ImageIndex = 0
          RightButton.Visible = True
          ShowHint = True
          TabOrder = 0
          OnChange = EPathExportChange
          OnRightButtonClick = EPathExportRightButtonClick
        end
      end
      object PPas: TPanel
        AlignWithMargins = True
        Left = 3
        Top = 158
        Width = 446
        Height = 21
        Align = alTop
        BevelOuter = bvNone
        TabOrder = 5
        object LPas: TLabel
          AlignWithMargins = True
          Left = 3
          Top = 3
          Width = 100
          Height = 15
          Align = alLeft
          AutoSize = False
          Caption = #1055#1072#1088#1086#1083#1100
        end
        object CBPas: TEdit
          Left = 106
          Top = 0
          Width = 340
          Height = 21
          Align = alClient
          TabOrder = 0
        end
      end
    end
    object TSAny: TTabSheet
      BorderWidth = 6
      Caption = #1056#1072#1079#1085#1086#1077
      ImageIndex = 2
      object CBModifyKod: TCheckBox
        Left = 0
        Top = 0
        Width = 452
        Height = 17
        Align = alTop
        Caption = #1056#1072#1079#1088#1077#1096#1080#1090#1100' '#1080#1079#1084#1077#1085#1077#1085#1080#1077' '#1089#1087#1080#1089#1082#1086#1074' '#1090#1072#1073#1083#1080#1094', '#1089#1090#1088#1086#1082' '#1080' '#1089#1090#1086#1083#1073#1094#1086#1074
        TabOrder = 0
        OnClick = CBModifyKodClick
        OnKeyPress = CBModifyKodKeyPress
      end
    end
  end
end
