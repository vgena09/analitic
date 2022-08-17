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
{===========================   СОЗДАНИЕ ФОРМЫ   ===============================}
{==============================================================================}
procedure TFINFO.FormCreate(Sender: TObject);
begin
    {Восстанавливаем настройки из Ini}
    LoadFormIni(Self, [fspPosition]);

    {Инициализация интерфейса}
    BtnClose1.ModalResult := mrClose;
    TreeView.Items.Clear;
    TreeView.ShowRoot := false;
    TreeView.ReadOnly := true;
    IsStop            := false; // Признак останова для внешней функции
end;


{==============================================================================}
{=============================   ПОКАЗ ФОРМЫ   ================================}
{==============================================================================}
procedure TFINFO.FormShow(Sender: TObject);
begin
    {На случай задержки при обращении к окну}
    Repaint;
end;


{==============================================================================}
{===========================   ЗАКРЫТИЕ ФОРМЫ   ===============================}
{==============================================================================}
procedure TFINFO.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {Сохраняем настройки в Ini}
    SaveFormIni(Self, [fspPosition]);
end;


{==============================================================================}
{==============    ДОБАВЛЯЕТ В TREEVIEW ИНФОРМАЦИОННУЮ СТРОКУ    ==============}
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

    {Прокрутка окна вниз}
    ScrollMessage.Msg        := WM_VScroll;
    ScrollMessage.ScrollCode := sb_LineDown;
    ScrollMessage.Pos        := 0;
    TreeView.Dispatch(ScrollMessage);

    {На случай длительной обработки}
    Application.ProcessMessages;
end;


{==============================================================================}
{==================    ACTION: ОТЧЕТ В ТЕКСТОВЫЙ РЕДАКТОР    ==================}
{==============================================================================}
procedure TFINFO.AReportExecute(Sender: TObject);
var SPath: String;
begin
    SPath := TempFileName(PATH_WORK_TEMP+'Отчет', 'txt');
    If Not ForceDirectories(ExtractFilePath(SPath)) then begin ErrMsg('Ошибка создания каталога: '+ExtractFilePath(SPath)); Exit; end;
    TreeView.SaveToFile(SPath);
    StartAssociatedExe(SPath, SW_SHOWNORMAL);
end;


{==============================================================================}
{==============    ACTION: ПРИЗНАК ОСТАНОВА ДЛЯ ВНЕШНЕЙ ФУНКЦИИ    ============}
{==============================================================================}
procedure TFINFO.ACloseExecute(Sender: TObject);
begin
    IsStop:=true;
end;


end.
