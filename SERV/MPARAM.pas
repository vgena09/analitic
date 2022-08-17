unit MPARAM;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.Math,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Menus,
  Vcl.ComCtrls, Vcl.ToolWin, Vcl.ExtCtrls, Vcl.ActnList, Vcl.Buttons,
  FunType,
  MAIN;

type
  TFPARAM = class(TForm)
    AList: TActionList;
    AOpenAs: TAction;
    ASave: TAction;
    OpenDlg: TOpenDialog;
    SaveDlg: TSaveDialog;
    AOpen: TAction;
    AClear: TAction;
    AAdd: TAction;
    ADel: TAction;
    PEdit: TPanel;
    TView: TTreeView;
    PColor: TPanel;
    Panel2: TPanel;
    PFormula: TPanel;
    EFormula: TEdit;
    BtnFormula: TBitBtn;
    Panel1: TPanel;
    CBColor: TComboBox;
    LColor: TLabel;
    LFormula: TLabel;
    PMenu: TPopupMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    AFormula: TAction;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);

    {������:  MPARAM_ACTION}
    procedure ASaveExecute(Sender: TObject);
    procedure AOpenAsExecute(Sender: TObject);
    procedure AOpenExecute(Sender: TObject);
    procedure AClearExecute(Sender: TObject);
    procedure AAddExecute(Sender: TObject);
    procedure ADelExecute(Sender: TObject);

    {������:  MPARAM_TVIEW}
    procedure FormShow(Sender: TObject);
    procedure TViewDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure TViewDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure TViewChange(Sender: TObject; Node: TTreeNode);
    procedure TViewCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure TViewDblClick(Sender: TObject);
    procedure EFormulaDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
    procedure EFormulaDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure CBColorChange(Sender: TObject);
    procedure EFormulaChange(Sender: TObject);
    procedure PMenuPopup(Sender: TObject);
    procedure AFormulaExecute(Sender: TObject);
    procedure TViewEdited(Sender: TObject; Node: TTreeNode; var S: string);
    procedure EFormulaMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure EFormulaMouseLeave(Sender: TObject);
    procedure EFormulaDblClick(Sender: TObject);
    procedure TViewEndDrag(Sender, Target: TObject; X, Y: Integer);

  private
    FFMAIN   : TFMAIN;

    {������:  MPARAM_ACTION}
    procedure EnablAction;

    {������:  MPARAM_IO}
    function  ParamWrite(const SPath: String): Boolean;
    function  ParamRead(const SPath: String): Boolean;
    procedure ParamClear;
    procedure TViewRefresh(const ISel: Integer);
  public
    LParam : TListParam;
  end;

const
  CONFIG_EXT    = 'prm';     // ���������� ������ ����������
  ITEM_MAX      = 200;       // ������������ ����� ����������
  SPR_SAV_STR   = '|';       // ������-��������� ��� ���������� ����������

var
  FPARAM  : TFPARAM;

implementation

uses FunConst, FunSys, FunVcl, FunText, FunIni, FunTree, FunInfo,
     MKOD;

{$R *.dfm}
{$INCLUDE MPARAM_ACTION}
{$INCLUDE MPARAM_TVIEW}
{$INCLUDE MPARAM_EDIT}
{$INCLUDE MPARAM_IO}


{==============================================================================}
{===========================   �������� �����   ===============================}
{==============================================================================}
procedure TFPARAM.FormCreate(Sender: TObject);
begin
    {�������������}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.FFPARAM := TForm(Self);
    SetLength(LParam, 0);

    TView.DragMode      := dmAutomatic;
    TView.Images        := FMAIN.ImgLstSys;
    TView.ToolTips      := false;
    TView.ShowRoot      := false;
    TView.HideSelection := false;
    TView.MultiSelect   := false;
    TView.ReadOnly      := false;

    EFormula.ShowHint   := true;

    CBColor.Style       := csDropDownList;
    CBColor.Items.Add('�� ��������');
    CBColor.Items.Add('����, > 0');
    CBColor.Items.Add('����');
    CBColor.Items.Add('��������, < 0');
    CBColor.Items.Add('��������');

    SaveDlg.DefaultExt  := CONFIG_EXT;
    SaveDlg.Filter      := '����� ���������� (*.'+CONFIG_EXT+')|*.'+CONFIG_EXT+'|��� ����� (*.*)|*.*';
    SaveDlg.Title       := '���������� ����������';
    OpenDlg.DefaultExt  := CONFIG_EXT;
    OpenDlg.Filter      := '����� ���������� (*.'+CONFIG_EXT+')|*.'+CONFIG_EXT+'|��� ����� (*.*)|*.*';
    OpenDlg.Title       := '�������� ����������';
end;


{==============================================================================}
{===========================   ����� �����   ===============================}
{==============================================================================}
procedure TFPARAM.FormShow(Sender: TObject);
begin
    {�������� ��������������}
    ParamRead('');
    LoadTreeSelect(@TView, INI_SET, INI_SET_PARAM_SELECT_ITEM);
    TViewChange(TView, TView.Selected);
end;


{==============================================================================}
{==========================   �������� �����   ================================}
{==============================================================================}
procedure TFPARAM.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {��������� ���������}
    ParamWrite('');
    SaveTreeSelect(@TView, INI_SET, INI_SET_PARAM_SELECT_ITEM);
    {����������� ������}
    ParamClear;
end;


{==============================================================================}
{=============================   ����� ����   =================================}
{==============================================================================}
procedure TFPARAM.PMenuPopup(Sender: TObject);
begin EnablAction; end;


end.
