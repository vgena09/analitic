unit MEXPORT;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Dialogs, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.FileCtrl,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ActnList, Vcl.Buttons,
  MAIN;

type
  TFEXPORT = class(TForm)
    ActionList: TActionList;
    AExport: TAction;
    ACancel: TAction;
    PPeriod: TPanel;
    LPeriod: TLabel;
    CBYear: TComboBox;
    CBMonth: TComboBox;
    PFolder: TPanel;
    LFolder: TLabel;
    PComment: TPanel;
    LComment: TLabel;
    CBComment: TMemo;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    PBottom: TPanel;
    Label1: TLabel;
    Panel4: TPanel;
    BtnCancel: TBitBtn;
    BtnOk: TBitBtn;
    CBFolder: TButtonedEdit;
    procedure FormCreate(Sender: TObject);
    procedure AExportExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure CBYearChange(Sender: TObject);
    procedure CBMonthChange(Sender: TObject);
    procedure CBFolderChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CBFolderRightButtonClick(Sender: TObject);
  private
    FFMAIN : TFMAIN;
    PasRar : String;

    procedure EnablAction;
    function  IsExistsAllRegions: Boolean;
  public
    function  Execute: Boolean;
  end;

var
  FEXPORT: TFEXPORT;

implementation

uses FunConst, FunText, FunDay, FunSys, FunInfo, FunVcl, FunFiles, FunBD, FunIni;

{$R *.dfm}


{==============================================================================}
{============================   СОЗДАНИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFEXPORT.FormCreate(Sender: TObject);
var I : Integer;
    B : Boolean;
begin
    {Инициализация}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));

    {Восстанавливаем настройки из Ini}
    LoadFormIni(Self, [fspPosition]);

    {Инициализация списка лет}
    For I:=CurrentYear-3 to CurrentYear do CBYear.Items.Add(IntToStr(I));

    {Инициализация списка месяцев}
    CBMonth.Style := csDropDownList;
    For I:=Low(MonthList) to High(MonthList) do CBMonth.Items.Add(MonthList[I]);

    {Читаем путь экспорта данных из файла местных настроек}
    CBFolder.Text := ReadLocalString(INI_SET, INI_SET_EXPORT_PATH, '');

    {Читаем пароль из файла глобальных настроек}
    PasRar := ReadGlobalString(INI_SET, INI_SET_PAS_IMPORT_EXPORT, '');

    {Доступность комментария}
    B := Not (Pos('-', PATH_WORK) > 0);
    LComment.Enabled  := B;
    CBComment.Enabled := B;
end;


{==============================================================================}
{============================   ЗАКРЫТИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFEXPORT.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {Сохраняем настройки в Ini}
    SaveFormIni(Self, [fspPosition]);
end;


{==============================================================================}
{============================   ВНЕШНЯЯ ФУНКЦИЯ    ============================}
{==============================================================================}
function TFEXPORT.Execute: Boolean;
var SYear, SMonth: String;
begin
    {Инициализация}
    Result:=false;

    {Определяем период по умолчанию}
    If Not NullPeriod(SYear, SMonth) then Exit;
    CBYear.Text  := SYear;
    CBMonth.ItemIndex := CBMonth.Items.IndexOf(SMonth);

    {Доступность Action}
    EnablAction;

    {Показ формы}
    Result := (ShowModal = mrOk);
end;


{==============================================================================}
{==========================   ДОСТУПНОСТЬ ACTION   ============================}
{==============================================================================}
procedure TFEXPORT.EnablAction;
var B1, B2 : Boolean;
    IYear  : Integer;
begin
    {Инициализация}
    B1 := DirectoryExists(PATH_BD_DATA+CBYear.Text+'\'+CBMonth.Text);
    B2 := DirectoryExists(CBFolder.Text);
    AExport.Enabled := B1 and B2;

    {Проверяем период}
    If IsIntegerStr(CBYear.Text) then begin
       IYear :=StrToInt(CBYear.Text);
       If (IYear>2100) or (IYear<1900) then B1 := false;
    end else B1:=false;

    If B1 then CBYear.Font.Color := clBlack
          else CBYear.Font.Color := clRed;
    CBMonth.Font.Color := CBYear.Font.Color;

    {Проверяем каталог}
    If B2 then CBFolder.Font.Color := clBlack
          else CBFolder.Font.Color := clRed;
end;


{==============================================================================}
{============================   ACTION: ЭКСПОРТ   =============================}
{==============================================================================}
procedure TFEXPORT.AExportExecute(Sender: TObject);
var SRar     : String;
    SDect    : String;
    SComment : String;
    SKey     : String;
begin
    {Предупреждение если не все регионы}
    If Not IsExistsAllRegions then Exit;

    {Инициализация}
    Hide;
    FFMAIN.Repaint; 
    SComment := '';
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Экспорт данных: сжатие';
    try
       {Файл архиватора}
       SRar   := GetProgramAssociation('.rar');
       If Not FileExists(SRar) then begin ErrMsg('Не обнаружена программа для сжатия файлов в RAR-архив!'); Exit; end;

       {Файл результата}
       SDect:=CBFolder.Text;
       If SDect[Length(SDect)] <> '\' then SDect:=SDect+'\';
       SDect:=SDect+'Статистика_'+CBYear.Text+'_'+CBMonth.Text;
       If SMAIN_REGION <> '' then SDect:=SDect+'_'+SMAIN_REGION;
       SDect:=SDect+'.rar';

       {Удаляем предыдущий файл результата}
       If FileExists(SDect) then begin
          If MessageDlg('Файл с аналогичным именем уже существует.'+CH_NEW+'Удалить данный файл?',
                        mtWarning, [mbYes, mbNo], 0)<>mrYes then Exit;
          DeleteFile(SDect);
       end;

       {Сохраняем комментарий если он есть}
       CBComment.Lines.Insert(0, ID_VERSION_PROG+GetProgFileVersion);
       If Trim(CBComment.Text) <> '' then begin
          SComment := PATH_WORK_TEMP+TEMP_COMMENT;
          SComment := TempFileName(ExtractFilePath(SComment)+ExtractFileNameWithoutExt(SComment), 'txt');
          If Not ForceDirectories(ExtractFilePath(SComment)) then begin ErrMsg('Ошибка создания каталога: '+ExtractFilePath(SComment)); Exit; end;
          CBComment.Lines.SaveToFile(SComment);
       end;

       {Параметры сжатия}
       //  t   - протестировать архив
       //  ap  - задать путь внутри архива
       //  ep1 - исключить из путей базовый каталог
       //  dh  - открывать совместно используемые файлы
       //  hp  - пароль
       //  k   - заблокировать архив
       //  m   - метод сжатия
       SKey := 'a -t -dh -m5 -ep1 -ap'+CBYear.Text;
       If PasRar  <> '' then SKey := SKey + ' -hp"'+PasRar+'"';
       If SComment = '' then SKey := SKey + ' -k';
(*
       {Удалить}
       SRar:=ReplStr(SRar, '7zFM.exe', '7z.exe');
       SRar:=ReplStr(SRar, '7zG.exe',  '7z.exe');

       SKey := 'a -ap'+CBYear.Text;
       If PasRar <> '' then SKey := SKey + ' -p"'+PasRar+'"';
       {Сжатие результатов}
       If Not ExecAndWait(SRar, SKey+' '+'"'+SDect+'" "'+PATH_BD_DATA+CBYear.Text+'\'+CBMonth.Text+'"', SW_SHOWNORMAL) then begin
          ErrMsg('Ошибка создания файла экспорта!');
          If FileExists(SDect) then DeleteFile(SDect);
          Exit;
       end;
*)
       {Сжатие результатов}
       If Not ExecAndWait(SRar, SKey+' -- '+'"'+SDect+'" "'+PATH_BD_DATA+CBYear.Text+'\'+CBMonth.Text+'"', SW_SHOWNORMAL) then begin
          ErrMsg('Ошибка создания файла экспорта!');
          If FileExists(SDect) then DeleteFile(SDect);
          Exit;
       end;

       {Запись комментария}
       If SComment<>'' then begin
          SKey := 'c -m5 -k -z"'+SComment+'"';
          If PasRar  <> '' then SKey := SKey + ' -hp"'+PasRar+'"';
          ExecAndWait(SRar, SKey+' -- '+'"'+SDect+'"', SW_MINIMIZE);
       end;

       {Если файла результата нет, то ошибка}
       If FileExists(SDect) then begin
          ForegroundWindow;
          MessageDlg('Файл экспорта создан!'+CH_NEW+SDect, mtInformation, [mbOk], 0);
          StartAssociatedExe('"'+ExtractFilePath(SDect)+'"', SW_SHOWNORMAL);
       end else begin
          ErrMsg('Ошибка создания файла экспорта!');
          Exit;
       end;

    finally
       {Информатор}
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
       {Удаляем файл комментария}
       If SComment<>'' then If FileExists(SComment) then DeleteFile(SComment);
       {Возвращаемый результат}
       ModalResult := mrNo;
    end;

    ModalResult := mrOk;
end;


{==============================================================================}
{============================   ACTION: ОТМЕНА   ==============================}
{==============================================================================}
procedure TFEXPORT.ACancelExecute(Sender: TObject);
begin
    ModalResult := mrCancel;
end;


{==============================================================================}
{========================   СОБЫТИЕ: ИЗМЕНЕНИЕ ГОДА   =========================}
{==============================================================================}
procedure TFEXPORT.CBYearChange(Sender: TObject);
begin
    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{======================   СОБЫТИЕ: ИЗМЕНЕНИЕ МЕСЯЦА   =========================}
{==============================================================================}
procedure TFEXPORT.CBMonthChange(Sender: TObject);
begin
    {Доступность Action}
    EnablAction;
end;

{==============================================================================}
{====================   СОБЫТИЕ: ИЗМЕНЕНИЕ ДИРЕКТОРИЯ   =======================}
{==============================================================================}
procedure TFEXPORT.CBFolderChange(Sender: TObject);
begin
    {Запоминаем путь экспорта данных}
    WriteLocalString(INI_SET, INI_SET_EXPORT_PATH, CBFolder.Text);
    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{======================   СОБЫТИЕ: ВЫБОР ДИРЕКТОРИЯ   =========================}
{==============================================================================}
procedure TFEXPORT.CBFolderRightButtonClick(Sender: TObject);
var SPath : String;
begin
    SPath := CBFolder.Text;
    If SelectDirectory('Каталог экспорта данных', '', SPath) then begin
       If SPath[Length(SPath)] <> '\' then SPath:=SPath+'\';
       CBFolder.Text := SPath;
    end;
end;

{==============================================================================}
{=========================   ВСЕ ЛИ РЕГИОНЫ ЕСТЬ   ============================}
{==============================================================================}
function TFEXPORT.IsExistsAllRegions: Boolean;
const SEPARATOR = '; ';
var LReg : TStringList;
    IReg : Integer;
    SErr : String;
begin
    {Инициализация}
    Result := false;
    SErr   := '';
    LReg   := TStringList.Create;
    try
       {Проверяем наличие файлов}
       If Not FFMAIN.ListRegionsRangeDown(@LReg, SMAIN_REGION, CBYear.Text, CBMonth.Text) then Exit;
       For IReg := 0 to LReg.Count -1 do begin
           If Not FileExists(FFMAIN.PathRegion(CBYear.Text, CBMonth.Text, LReg[IReg])) then begin
              SErr := SErr + LReg[IReg] + SEPARATOR;
              If Length(SErr) > 200 then begin
                 Delete(SErr, 200, Length(SErr));
                 SErr := SErr + '...'+SEPARATOR;
                 Break;
              end;
           end;
       end;
    finally
       LReg.Free;
    end;
    If SErr <> '' then Delete(SErr, Length(SErr) - Length(SEPARATOR) + 1, Length(SEPARATOR));

    {Подтверждение}
    If SErr <> '' then begin
       If MessageDlg('Отсутствуют следующие отчеты:'+CH_NEW+SErr+CH_NEW+CH_NEW+
                     'Подтвердите операцию.', mtWarning, [mbOk, mbCancel], 0) <> mrOk then Exit;
    end;

    {Результат}
    Result := true;
end;

end.
