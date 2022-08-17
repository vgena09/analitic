unit MBLANK;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.FileCtrl,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.CheckLst, Vcl.ActnList, Vcl.Buttons,
  IdGlobal, IdGlobalProtocols,
  MAIN;

type
  TFBLANK = class(TForm)
    CBView: TCheckListBox;
    LCaption: TLabel;
    PSpace1: TPanel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    BtnCancel: TBitBtn;
    AList: TActionList;
    AOk: TAction;
    ACancel: TAction;
    Panel4: TPanel;
    BtnOk: TBitBtn;
    Panel5: TPanel;
    LFolder: TLabel;
    CBFolder: TButtonedEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure CBFolderChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure CBViewClick(Sender: TObject);
    procedure CBViewKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure AOkExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure CBFolderRightButtonClick(Sender: TObject);
  private
    FFMAIN    : TFMAIN;
    DatPeriod : TDate;

    procedure EnablAction;
  public

  end;

var
  FBLANK: TFBLANK;

implementation

uses FunConst, FunText, FunSys, FunVcl, FunBD, FunDay, FunExcel, FunIni;

{$R *.dfm}


{==============================================================================}
{============================   СОЗДАНИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFBLANK.FormCreate(Sender: TObject);
var LForma_ : TLForma;
    I       : Integer;
begin
    {Инициализация}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Запись бланка отчета для регионов';
    DatPeriod := DateTimeCorrect(Now, 0, 1, 0, 0, 0);

    {Формируем список отчетов}
    CBView.Items.Clear;
    try
       If Not FFMAIN.FindForms(@LForma_, false, true, -1, -1, '', '', '', '', false, DatPeriod) then begin
          ErrMsg('Ошибка формирования списка отчетов!');
          Exit;
       end;
       For I:=Low(LForma_) to High(LForma_) do begin
          If Not LForma_[I].FExport   then Continue;
          If LForma_[I].FCaption = '' then Continue;
          CBView.Items.Add(LForma_[I].FCaption);
       end;
    finally
       SetLength(LForma_, 0);
    end;
    {Читаем путь экспорта данных из файла местных настроек}
    CBFolder.Text := ReadLocalString(INI_SET, INI_SET_EXPORT_PATH, '');

    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{============================   АКТИВАЦИЯ ФОРМЫ    ============================}
{==============================================================================}
procedure TFBLANK.FormActivate(Sender: TObject);
begin
    CBView.SetFocus; 
end;


{==============================================================================}
{=============================   СКРЫТИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFBLANK.FormHide(Sender: TObject);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;


{==============================================================================}
{========================   ACTION: ЗАПИСЬ ОТЧЕТОВ    =========================}
{==============================================================================}
procedure TFBLANK.AOkExecute(Sender: TObject);
var LForma_ : TLForma;
    SSrc, STemp, SDect: String;
    I : Integer;
    E : Variant;
begin
    {Инициализация}
    ModalResult:=mrCancel;

    {Ищем отмеченные отчеты}
    For I:=0 to CBView.Count-1 do begin
       If Not CBView.Checked[I] then Continue;
       FFMAIN.Repaint; 

       try {Ищем отмеченную форму}
           If Not FFMAIN.FindForms(@LForma_, true, true, -1, -1, '', CBView.Items[I], '', '', false, DatPeriod) then begin ErrMsg('Не найдена запись отчета: '+CBView.Items[I]); Continue; end;
           {Файл-испочник}
           SSrc:=FFMAIN.PathFileForm(LForma_[0].SBankPref, LForma_[0].FFile);
           If Not FileExists(SSrc) then begin ErrMsg('Не найден файл отчета: '+SSrc); Continue; end;
       finally
           SetLength(LForma_, 0);
       end;

       {Файл-результат}
       SDect:=CBFolder.Text;
       If SDect[Length(SDect)] <> '\' then SDect:=SDect+'\';
       SDect:=SDect+ExtractFileName(SSrc);

       {Удаляем предыдущий файл результата}
       If FileExists(SDect) then begin
          If MessageDlg('Файл с аналогичным именем уже существует:'+CH_NEW+SDect+CH_NEW+'Удалить данный файл?',
                        mtWarning, [mbYes, mbNo], 0)<>mrYes then begin ModalResult:=mrNone; Continue; end;
          FFMAIN.Repaint; 
          DeleteFile(SDect);
       end;

       {Информатор}
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Сохранение бланка отчета: '+CBView.Items[I];
       FFMAIN.Repaint; 

       {Временный файл}
       STemp := CopyFileToTemp(SSrc);
       If STemp='' then begin ErrMsg('Ошибка создания временного файла!'); Continue; end;

       {Корректируем отчет}
       If Not CreateExcel then begin ErrMsg('Ошибка открытия MS Excel!'); Continue; end;
       try If Not OpenXls(STemp, true) then begin ErrMsg('Ошибка открытия файла: '+CH_NEW+STemp); Continue; end;
           try
              E:=GetExcelApplicationID;
              E.DisplayAlerts:=false;
              E.ActiveWorkbook.WorkSheets[1].Range[CELL_FULL].Value:='FULL';
              If Not SaveXls then begin ErrMsg('Ошибка сохранения временного файла: '+CH_NEW+STemp); Continue; end;
           finally
              If Not CloseXls then ErrMsg('Ошибка закрытия импортируемого отчета!');
           end;
       finally
          If Not CloseExcel then ErrMsg('Ошибка закрытия Excel!');
          E := Unassigned;
       end;

       {Копируем файл}
       If Not ForceDirectories(ExtractFilePath(SDect)) then begin ErrMsg('Ошибка создания каталога: '+ExtractFilePath(SDect)); Continue; end;
       If Not CopyFileTo(STemp, SDect)                 then begin ErrMsg('Ошибка создания файла отчета: '+SDect); Continue; end;

       {Удаляем временный файл}
       DeleteFile(STemp);
    end;

    {Открываем папку}
    StartAssociatedExe('"'+ExtractFilePath(SDect)+'"', SW_SHOWNORMAL);

    {Возвращаемый результат}
    ModalResult:=mrOk;
end;



{==============================================================================}
{===========================   ACTION: ОТМЕНА    ==============================}
{==============================================================================}
procedure TFBLANK.ACancelExecute(Sender: TObject);
begin
    ModalResult := mrCancel;
end;


{==============================================================================}
{==========================   ДОСТУПНОСТЬ ACTION   ============================}
{==============================================================================}
procedure TFBLANK.EnablAction;
var I : Integer;
    B : Boolean;
begin
    {Должен быть хоть один выбор}
    B := false;
    For I:=0 to CBView.Count-1 do begin
       If CBView.Checked[I] then B:=true;
    end;

    {Каталог должен быть задан}
    If Not DirectoryExists(CBFolder.Text) then B:=false;

    {Устанавливаем доступность}
    AOk.Enabled := B;
end;


{==============================================================================}
{====================   СОБЫТИЕ: ИЗМЕНЕНИЕ ДИРЕКТОРИЯ   =======================}
{==============================================================================}
procedure TFBLANK.CBFolderChange(Sender: TObject);
begin
    {Запоминаем путь экспорта данных}
    WriteLocalString(INI_SET, INI_SET_EXPORT_PATH, CBFolder.Text);
    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{======================   СОБЫТИЕ: ВЫБОР ДИРЕКТОРИЯ   =========================}
{==============================================================================}
procedure TFBLANK.CBFolderRightButtonClick(Sender: TObject);
var SPath : String;
begin
    SPath := CBFolder.Text;
    If SelectDirectory('Место сохранения бланков', '', SPath) then begin
       If SPath[Length(SPath)] <> '\' then SPath:=SPath+'\';
       CBFolder.Text := SPath;
    end;
end;

{==============================================================================}
{==========================   СОБЫТИЕ: ON_CLICK   =============================}
{==============================================================================}
procedure TFBLANK.CBViewClick(Sender: TObject);
begin
    {Доступность Action}
    EnablAction;
end;

procedure TFBLANK.CBViewKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    CBViewClick(Sender);
end;

end.
