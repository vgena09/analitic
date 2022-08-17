unit MINFO;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ActnList,
  Vcl.ImgList, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  FunType;

type
  TFINFO = class(TForm)
    PBottom: TPanel;
    ImLst: TImageList;
    BtnClose1: TBitBtn;
    ActionList1: TActionList;
    AReport: TAction;
    AClose: TAction;
    Panel1: TPanel;
    BtnReport: TBitBtn;
    TreeView: TTreeView;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure AReportExecute(Sender: TObject);
    procedure ACloseExecute(Sender: TObject);
  private

  public
    IsStop : Boolean;
    procedure AddInfo(const Msg: String; const IndIco: Integer);
  end;

const
  ICO_INFO = 0;
  ICO_WARN = 1;
  ICO_OK   = 2;
  ICO_ERR  = 3;

var
  FINFO: TFINFO;

implementation

uses FunConst, Main, FunSys, FunVcl, FunFiles, FunIni;

{$R *.dfm}


{==============================================================================}
{===========================   �������� �����   ===============================}
{==============================================================================}
procedure TFINFO.FormCreate(Sender: TObject);
begin
    {��������������� ��������� �� Ini}
    LoadFormIni(Self, [fspPosition]);

    {������������� ����������}
    BtnClose1.ModalResult := mrClose;
    TreeView.Items.Clear;
    TreeView.ShowRoot := false;
    TreeView.ReadOnly := true;
    IsStop            := false; // ������� �������� ��� ������� �������
end;


{==============================================================================}
{=============================   ����� �����   ================================}
{==============================================================================}
procedure TFINFO.FormShow(Sender: TObject);
begin
    {�� ������ �������� ��� ��������� � ����}
    Repaint;
end;


{==============================================================================}
{===========================   �������� �����   ===============================}
{==============================================================================}
procedure TFINFO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {��������� ��������� � Ini}
    SaveFormIni(Self, [fspPosition]);
end;


{==============================================================================}
{==============    ��������� � TREEVIEW �������������� ������    ==============}
{==============================================================================}
procedure TFINFO.AddInfo(const Msg: String; const IndIco: Integer);
var Item          : TTreeNode;
    ScrollMessage : TWMVScroll;
begin
    TreeView.Items.BeginUpdate;
    try Item            := TreeView.Items.Add(nil, Msg);
        Item.ImageIndex := IndIco;
    finally
       TreeView.Items.EndUpdate;
    end;

    {��������� ���� ����}
    ScrollMessage.Msg        := WM_VScroll;
    ScrollMessage.ScrollCode := sb_LineDown;
    ScrollMessage.Pos        := 0;
    TreeView.Dispatch(ScrollMessage);

    {�� ������ ���������� ���������}
    Application.ProcessMessages;
end;


{==============================================================================}
{==================    ACTION: ����� � ��������� ��������    ==================}
{==============================================================================}
procedure TFINFO.AReportExecute(Sender: TObject);
var SPath: String;
begin
    SPath := TempFileName(PATH_WORK_TEMP+'�����', 'txt');
    If Not ForceDirectories(ExtractFilePath(SPath)) then begin ErrMsg('������ �������� ��������: '+ExtractFilePath(SPath)); Exit; end;
    TreeView.SaveToFile(SPath);
    StartAssociatedExe(SPath, SW_SHOWNORMAL);
end;


{==============================================================================}
{==============    ACTION: ������� �������� ��� ������� �������    ============}
{==============================================================================}
procedure TFINFO.ACloseExecute(Sender: TObject);
begin
    IsStop:=true;
end;


end.
