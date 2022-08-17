object FPARAM: TFPARAM
  Left = 1059
  Top = 354
  AlphaBlend = True
  AlphaBlendValue = 100
  BorderStyle = bsSizeToolWin
  Caption = #1055#1072#1088#1072#1084#1077#1090#1088#1099' '#1072#1085#1072#1083#1080#1090#1080#1082#1080
  ClientHeight = 277
  ClientWidth = 460
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = False
  PopupMenu = PMenu
  Position = poDefault
  OnClose = FormClose
  OnCreate = FormCreate
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object PEdit: TPanel
    Left = 0
    Top = 213
    Width = 460
    Height = 64
    Align = alBottom
    AutoSize = True
    BevelOuter = bvNone
    TabOrder = 0
    object PColor: TPanel
      Left = 0
      Top = 41
      Width = 460
      Height = 23
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 0
      object LColor: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 65
        Height = 17
        Hint = #1059#1089#1083#1086#1074#1080#1077' '#1074#1099#1076#1077#1083#1077#1085#1080#1103' '#1079#1085#1072#1095#1077#1085#1080#1103' '#1079#1077#1083#1077#1085#1099#1084' '#1094#1074#1077#1090#1086#1084
        Align = alLeft
        AutoSize = False
        Caption = #1042#1099#1076#1077#1083#1077#1085#1080#1077
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clGreen
        Font.Height = -11
        Font.Name = 'MS Sans Serif'
        Font.Style = []
        ParentFont = False
        ParentShowHint = False
        ShowHint = True
      end
      object CBColor: TComboBox
        Left = 71
        Top = 0
        Width = 389
        Height = 21
        Align = alClient
        Style = csDropDownList
        TabOrder = 0
        OnChange = CBColorChange
      end
    end
    object Panel2: TPanel
      Left = 0
      Top = 32
      Width = 460
      Height = 9
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 1
    end
    object PFormula: TPanel
      Left = 0
      Top = 9
      Width = 460
      Height = 23
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 2
      object LFormula: TLabel
        AlignWithMargins = True
        Left = 3
        Top = 3
        Width = 65
        Height = 17
        Align = alLeft
        AutoSize = False
        Caption = #1060#1086#1088#1084#1091#1083#1072
      end
      object EFormula: TEdit
        Left = 71
        Top = 0
        Width = 353
        Height = 23
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 0
        OnChange = EFormulaChange
        OnDblClick = EFormulaDblClick
        OnDragDrop = EFormulaDragDrop
        OnDragOver = EFormulaDragOver
        OnMouseLeave = EFormulaMouseLeave
        OnMouseMove = EFormulaMouseMove
        ExplicitHeight = 24
      end
      object BtnFormula: TBitBtn
        Left = 424
        Top = 0
        Width = 36
        Height = 23
        Action = AFormula
        Align = alRight
        Caption = '...'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
      end
    end
    object Panel1: TPanel
      Left = 0
      Top = 0
      Width = 460
      Height = 9
      Align = alBottom
      BevelOuter = bvNone
      TabOrder = 3
    end
  end
  object TView: TTreeView
    Left = 0
    Top = 0
    Width = 460
    Height = 213
    Align = alClient
    BorderStyle = bsNone
    HideSelection = False
    Images = FMAIN.ImgLstSys
    Indent = 19
    ShowRoot = False
    TabOrder = 1
    ToolTips = False
    OnChange = TViewChange
    OnCustomDrawItem = TViewCustomDrawItem
    OnDblClick = TViewDblClick
    OnDragDrop = TViewDragDrop
    OnDragOver = TViewDragOver
    OnEdited = TViewEdited
    OnEndDrag = TViewEndDrag
    Items.NodeData = {
      03030000002A0000000000000000000000FFFFFFFFFFFFFFFF00000000000000
      000000000001063F043004320430043F0432042E0000000000000000000000FF
      FFFFFFFFFFFFFF0000000000000000000000000108320430043F04320430043F
      0432043F042C0000000000000000000000FFFFFFFFFFFFFFFF00000000000000
      0000000000010730043F04320430043F0432043004}
  end
  object AList: TActionList
    Images = FMAIN.ImgLstSys
    Left = 96
    Top = 96
    object AClear: TAction
      Caption = #1054#1095#1080#1089#1090#1080#1090#1100' '#1089#1087#1080#1089#1086#1082
      ImageIndex = 4
      OnExecute = AClearExecute
    end
    object AAdd: TAction
      Caption = #1044#1086#1073#1072#1074#1080#1090#1100' '#1087#1072#1088#1072#1084#1077#1090#1088
      ImageIndex = 0
      ShortCut = 16429
      OnExecute = AAddExecute
    end
    object ADel: TAction
      Caption = #1059#1076#1072#1083#1080#1090#1100' '#1087#1072#1088#1072#1084#1077#1090#1088
      ImageIndex = 1
      ShortCut = 16430
      OnExecute = ADelExecute
    end
    object ASave: TAction
      Caption = #1057#1086#1093#1088#1072#1085#1080#1090#1100' '#1089#1087#1080#1089#1086#1082' ...'
      ImageIndex = 3
      ShortCut = 113
      OnExecute = ASaveExecute
    end
    object AOpen: TAction
      Caption = #1054#1090#1082#1088#1099#1090#1100' '#1089#1087#1080#1089#1086#1082' ...'
      ImageIndex = 2
      ShortCut = 114
      OnExecute = AOpenExecute
    end
    object AOpenAs: TAction
      Caption = #1054#1090#1082#1088#1099#1090#1100' '#1089#1074#1086#1081' '#1089#1087#1080#1089#1086#1082' ...'
      ImageIndex = 2
      ShortCut = 115
      OnExecute = AOpenAsExecute
    end
    object AFormula: TAction
      Caption = '...'
      Hint = #1056#1077#1076#1072#1082#1090#1086#1088' '#1092#1086#1088#1084#1091#1083#1099' ...'
      OnExecute = AFormulaExecute
    end
  end
  object OpenDlg: TOpenDialog
    DefaultExt = 'set1'
    Filter = #1060#1072#1081#1083#1099' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080'|*.set1|'#1042#1089#1077' '#1092#1072#1081#1083#1099'|*.*'
    Title = #1060#1072#1081#1083' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080
    Left = 128
    Top = 96
  end
  object SaveDlg: TSaveDialog
    DefaultExt = 'set1'
    Filter = #1060#1072#1081#1083#1099' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080'|*.set1|'#1042#1089#1077' '#1092#1072#1081#1083#1099'|*.*'
    Title = #1060#1072#1081#1083' '#1082#1086#1085#1092#1080#1075#1091#1088#1072#1094#1080#1080
    Left = 160
    Top = 96
  end
  object PMenu: TPopupMenu
    Images = FMAIN.ImgLstSys
    MenuAnimation = [maTopToBottom]
    OnPopup = PMenuPopup
    Left = 192
    Top = 96
    object N1: TMenuItem
      Action = AClear
    end
    object N2: TMenuItem
      Action = AAdd
    end
    object N3: TMenuItem
      Action = ADel
    end
    object N4: TMenuItem
      Caption = '-'
    end
    object N5: TMenuItem
      Action = ASave
    end
    object N6: TMenuItem
      Action = AOpen
    end
    object N7: TMenuItem
      Action = AOpenAs
    end
  end
end
