unit MIMPORT_RAR;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.StdCtrls, Vcl.Menus,
  Vcl.Dialogs, Vcl.ActnList, Vcl.CheckLst, Vcl.ExtCtrls, Vcl.Buttons,
  Data.DB, Data.Win.ADODB, IdGlobal,
  MAIN;

type
  TFIMPORT_RAR = class(TForm)
    ActionList: TActionList;
    AImport: TAction;
    ACancel: TAction;
    CBView: TCheckListBox;
    LComment: TLabel;
    CBComment: TMemo;
    PMenu: TPopupMenu;
    ASelectAll: TAction;
    AUnselectAll: TAction;
    N1: TMenuItem;
    N2: TMenuItem;
    LCaption: TLabel;
    CBSubReg: TCheckBox;
    CBVerify: TCheckBox;
    PBottom: TPanel;
    PSpaceComment: TPanel;
    Panel1: TPanel;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel2: TPanel;
    BtnOk: TBitBtn;
    BtnCancel: TBitBtn;
    PVersion: TPanel;
    procedure FormCreate(Sender: TObject);
    procedure EnablAction;
    procedure AImportExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure CBSubRegClick(Sender: TObject);
    procedure CBSubRegKeyPress(Sender: TObject; var Key: Char);
    procedure CBVerifyClick(Sender: TObject);
    procedure CBVerifyKeyPress(Sender: TObject; var Key: Char);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ASelectAllExecute(Sender: TObject);
    procedure AUnselectAllExecute(Sender: TObject);
    procedure CBViewClick(Sender: TObject);
    procedure CBViewKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
  private
    FFMAIN          : TFMAIN;
    STemp, STempDat : String;
    SYear, SMonth   : String;        // Импортируемый период
    DatPeriod       : TDate;
  public
    function  Execute(const SPath: String): Boolean;
  end;

var
  FIMPORT_RAR : TFIMPORT_RAR;
  PasRar      : String;

implementation

uses FunConst,
     FunText, FunSys, FunVcl, FunInfo, FunFiles,
     FunBD, FunDay, FunSum, FunVerify, FunIO, FunIni,
     MINFO;

{$R *.dfm}


{==============================================================================}
{============================   СОЗДАНИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFIMPORT_RAR.FormCreate(Sender: TObject);
begin
    {Инициализация}
    FFMAIN:= TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Импорт отчетов из архива';
    PVersion.Visible := false;

    {Восстанавливаем настройки из Ini}
    LoadFormIni(Self, [fspPosition]);

    {Читаем пароль}
    PasRar := ReadGlobalString(INI_SET, INI_SET_PAS_IMPORT_EXPORT, '');

    {Читаем признак импорта субрегионов, проверки отчетов}
    CBSubReg.Checked := ReadLocalBool(INI_SET, INI_SET_IMPORT_SUBREG,     true);
    CBVerify.Checked := ReadLocalBool(INI_SET, INI_SET_IMPORT_VERIFY_RAR, true);
end;


{==============================================================================}
{============================   ЗАКРЫТИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFIMPORT_RAR.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {Сохраняем настройки в Ini}
    SaveFormIni(Self, [fspPosition]);

    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;

{==============================================================================}
{============================   ВНЕШНЯЯ ФУНКЦИЯ    ============================}
{==============================================================================}
function TFIMPORT_RAR.Execute(const SPath: String): Boolean;
var Sr            : TSearchRec;
    STempCom      : String;
    S, SRar, SKey : String;
    SRegion       : String;
    IYear, IMonth : Integer;
    B             : Boolean;
begin
    {Инициализация}
    Result   := false;
    If Not FileExists(SPath) then begin ErrMsg('Файл данных не найден!'+CH_NEW+SPath); Exit; end;
    SYear    := '';
    SMonth   := '';

    STemp    := PATH_WORK_TEMP+PATH_TEMP_RAR;
    STempDat := STemp+'Импортируемые данные\';
    STempCom := STemp+'Комментарий.txt';
    If Not DelDir(STemp, true) then begin ErrMsg('Ошибка удаления предыдущих файлов!'+STemp); Exit; end;

    {Файл архиватора}
    SRar := GetProgramAssociation('.rar');
    If Not FileExists(SRar) then begin ErrMsg('Не обнаружена программа для распаковки файлов из RAR-архива!'); Exit; end;

    {Распаковка данных}
    SKey := 'x ';
    If PasRar <> '' then SKey := SKey + ' -p"'+PasRar+'"';
    If Not ExecAndWait(SRar, SKey+' -- '+'"'+SPath+'" "'+STempDat+'"', SW_SHOWNORMAL) then begin ErrMsg('Ошибка распаковки данных!'); Exit; end;
    try
       {Имеются ли извлеченные данные}
       B:=false;
       try
          If FindFirst(STempDat+'*', faDirectory, Sr) = 0 then begin
             repeat If (Sr.Name<>'.') and (Sr.Name<>'..') then begin B:=true; Break; end;
             until  FindNext(Sr) <> 0;
          end;
       finally FindClose(Sr);
       end;
       If Not B then Exit;

       {Извлечение комментария}
       SKey := 'cw ';
       If PasRar <> '' then SKey := SKey + ' -p"'+PasRar+'"';
       If Not ExecAndWait(SRar, SKey+' -- '+'"'+SPath+'" "'+STempCom+'"', SW_SHOWNORMAL) then begin ErrMsg('Ошибка распаковки данных!'); Exit; end;
       If FileExists(STempCom) then CBComment.Lines.LoadFromFile(STempCom) 
                               else CBComment.Lines.Clear;

       {Проверка версии программы отправителя}
       SKey := 'старая';
       If CBComment.Lines.Count > 0  then begin
          S := CBComment.Lines[0];
          If Pos(ID_VERSION_PROG, S) > 0 then begin
             SKey := Copy(S, Length(ID_VERSION_PROG)+1, Length(S));
             CBComment.Lines.Delete(0);  
          end else begin
             SKey := '';
          end;
       end;
       If SKey <> GetProgFileVersion then begin
          PVersion.Caption := 'Версия программы отправителя ['+SKey+'] отлична от Вашей ['+GetProgFileVersion+']';
          PVersion.Visible := true;
       end;

       {Показ комментария}
       If CBComment.Lines.Count > 0  then begin
          CBComment.Visible     := true;
          LComment.Visible      := true;
          PSpaceComment.Visible := true;
       end else begin
          CBComment.Visible     := false;
          LComment.Visible      := false;
          PSpaceComment.Visible := false;
       end;
      
       {Определяем импортируемый год - SYear}
       DatPeriod := 0;
       IYear     := 0;
       try
          If FindFirst(STempDat+'*', faDirectory, Sr) = 0 then begin
             repeat If IsIntegerStr(Sr.Name) then begin
                       SYear:=Sr.Name;
                       IYear:=StrToInt(SYear);
                       If (IYear>2100) or (IYear<1900) then SYear:='';
                       Break;
                    end;
             until  FindNext(Sr) <> 0;
          end;
       finally FindClose(Sr);
       end;
       If SYear='' then begin ErrMsg('Не определен год импортируемых данных!'); Exit; end;

       {Определяем импортируемый месяц - SMonth}
       IMonth:=0;
       try
          If FindFirst(STempDat+SYear+'\*', faDirectory, Sr) = 0 then begin
             repeat If MonthStrToInd(Sr.Name)>0 then begin
                       SMonth := Sr.Name;
                       IMonth := MonthStrToInd(Sr.Name);
                       If (IMonth<Low(MonthList)) or (IMonth>High(MonthList)) then SMonth:='';
                       Break;
                    end;
             until  FindNext(Sr) <> 0;
          end;
       finally FindClose(Sr);
       end;
       If SMonth='' then begin ErrMsg('Не определен месяц импортируемых данных!'); Exit; end;

       {Определяем период}
       DatPeriod:=IDatePeriod(IYear, IMonth);
       If DatPeriod=0 then begin ErrMsg('Недопустимый период!'); Exit; end;

       {Заполняем список импортируемых регионов}
       CBView.Items.Clear;
       try
          If FindFirst(STempDat+SYear+'\'+SMonth+'\*.dat', faDirectory, Sr) = 0 then begin
             repeat SRegion := ExtractFileNameWithoutExt(Sr.Name);
                    If FFMAIN.VerifyRegionName(SRegion) then CBView.Items.Add(SRegion);
             until  FindNext(Sr) <> 0;
          end;
       finally FindClose(Sr);
       end;

       {Доступность ACTION}
       ASelectAllExecute(nil);

       {Настройка предварительного просмотра}
       CBViewClick(nil);

       {Показ формы}
       Result := (ShowModal = mrOk);

    finally
       DelDir(STemp, true);
    end;
end;


{==============================================================================}
{==========================   ДОСТУПНОСТЬ ACTION   ============================}
{==============================================================================}
procedure TFIMPORT_RAR.EnablAction;
var BImport, BSel, BUnSel : Boolean;
    I, ISel               : Integer;
    B                     : Boolean;
begin
    {Инициализация}
    BImport := false;
    BSel    := false;
    BUnSel  := false;
    B       := Not CBSubReg.Checked;
    ISel    := 0;

    For I:=0 to CBView.Count - 1 do begin
       {Корректируем выделение и доступность субрегионов}
       If B then begin
          If CutModulStr(CBView.Items[I], SUB_RGN1, SUB_RGN2)<>'' then begin
             CBView.ItemEnabled[I] := false;
             CBView.Checked[I]     := false;
          end;
       end else begin
          CBView.ItemEnabled[I]    := true;
       end;

       {Доступность}
       If CBView.Checked[I] then begin
          BImport := true;
          BUnSel  := true;
          ISel    := ISel+1;
       end else begin
          BSel    := true;
       end;

    end;

    {Перерисовываем окно}
    LCaption.Caption := 'Выбрано '+IntToStr(ISel)+' рег. из '+IntToStr(CBView.Items.Count);
    CBView.Refresh;

    {Установка ACTION}
    AImport.Enabled      := BImport;
    ASelectAll.Enabled   := BSel;
    AUnSelectAll.Enabled := BUnSel;
end;


{==============================================================================}
{=============================   ACTION: ИМПОРТ   =============================}
{==============================================================================}
procedure TFIMPORT_RAR.AImportExecute(Sender: TObject);
var FFMAIN               : TFMAIN;
    FFINFO               : TFINFO;
    LRegion_             : TLRegion;
    LTab, LTabMain, LDat : TStringList;
    SReg                 : String;
    SReg1, SReg2         : String;
    SSrc, SDect          : String;
    CountReg             : Integer;     // Число успешно импортированных регионов
    IMainParent          : Integer;     // Родитель главного элемента
    SMainReg             : String;      // Регион главного элемента
    SMainMain            : String;      // Родитель главного элемента
    S                    : String;
    I, IReg              : Integer;
    IsError              : Boolean;
begin
    {Предупреждение при попытке импорта главного региона}
    I := CBView.Items.IndexOf(SMAIN_REGION);
    If  I > 0 then begin
       If CBView.Checked[I] then begin
          If MessageDlg('Выбран импорт отчета главного региона: ['+SMAIN_REGION+'].'+CH_NEW+
                        'Подтвердите операцию.', mtWarning, [mbOk, mbCancel], 0) <> mrOk then Exit;
       end;
    end;

    {Инициализация}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Импорт данных: инициализация';
    IsError     := false;
    CountReg    := 0;
    IMainParent := 999999999;
    SMainReg    := '';
    LTab        := TStringList.Create;
    LTabMain    := TStringList.Create;
    LDat        := TStringList.Create;

    {Скрываем окно импорта и обновляем основное окно перед длительной операцией}
    Hide;
    FFMAIN.Repaint;

    FFINFO:=TFINFO.Create(Self);
    try
       {Инициализация информатора}
       FFINFO.Caption := 'Импорт данных';
       FFINFO.AReport.Enabled := false;
       FFINFO.Show;

       {Импортируем данные регионов}
       For IReg:=0 to CBView.Items.Count-1 do begin
          {Не выделенные регионы пропускаем}
          If Not CBView.Checked[IReg] then Continue;

          {Название региона}
          SReg  := CBView.Items[IReg];
          SReg2 := CutModulStr(SReg,      SUB_RGN1, SUB_RGN2);
          SReg1 := ReplModulStr(SReg, '', SUB_RGN1, SUB_RGN2);

          {Информатор}
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Импорт данных: '+IntToStr(CountReg+1)+'. '+SReg;

          {Проверка названия региона}
          If FFMAIN.FindRegions(@LRegion_, true, -1, SReg1, -1, -1, DatPeriod) then begin
             {Корректируем главный регион}
             I:=LRegion_[0].FParent;
             If I<IMainParent then begin
                IMainParent := I;
                SMainReg    := SReg1;
             end;
          end else begin
             IsError:=true;
             FFINFO.AddInfo(SReg1+': Недопустимый регион', ICO_ERR);
             Continue;
          end;

          {Пути}
          SSrc  := STempDat+SYear+'\'+SMonth+'\'+SReg+'.dat';
          SDect := PATH_BD_DATA+SYear+'\'+SMonth+'\'+SReg+'.dat';

          {Список импортируемых таблиц}
          FFMAIN.GetListTabDat(@LTab, SSrc);
          If SMainReg = SReg1 then LTabMain.Text:=LTab.Text;
          If LTab.Count = 0 then begin
             IsError:=true;
             FFINFO.AddInfo(SReg+': Нет данных', ICO_ERR);
             Continue;
          end;

          {Переносим таблицы}
          For I:=0 to LTab.Count-1 do begin
             FFMAIN.GetTabVal(SSrc, LTab[I], @LDat);
             FFMAIN.SetTabVal(SDect, LTab[I], @LDat, true);
          end;

          {Информатор}
          Inc(CountReg);
          FFINFO.AddInfo(SReg+': Импортирован', ICO_OK);
       end;

       {Корректируем выделение}
       S := FFMAIN.CreateRegions(SYear, SMonth, SMainReg);
       WriteLocalString(INI_SELECT, INI_SELECT_PATH, S);

       {Попытка расчета данных с учетом импорта}
       If (Not IsError) and (SMainReg<>'') and (IMainParent>0) and (LTabMain.Count>0) then begin
          SMainMain:=FFMAIN.RegionsCounterToCaption(IMainParent);
          FFINFO.AddInfo(SMainMain+': расчет за '+SMonth+' '+SYear+' г.', ICO_INFO);
          For I:=0 to LTabMain.Count-1 do begin
             S:=CutLongStr(FFMAIN.TablesCounterToCaption(LTabMain[I]), 80);
             If SumTable(LTabMain[I], SYear, SMonth, SMainReg)
             then FFINFO.AddInfo('Расчет произведен: '+S, ICO_OK)
             else FFINFO.AddInfo('Не расчитано: '+S+'. Возможно не собраны все отчеты.', ICO_WARN);
          end;
       end else begin
          FFINFO.AddInfo('Расчет данных пропущен', ICO_INFO)
       end;

       {Проверка отчета}
       If CBVerify.Checked and (Not IsError) and (SMainReg<>'') and (IMainParent>0) and (LTabMain.Count>0)then begin
          VerifyTable(@LTabMain, SYear, SMonth, SMainReg, true);
       end else begin
          FFINFO.AddInfo('Проверка данных пропущена', ICO_INFO)
       end;
    finally
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
       SetLength(LRegion_, 0);
       LDat.Free;
       LTabMain.Free;
       LTab.Free;

       {Результаты импорта}
       If IsError then FFINFO.AddInfo('Импортировано только '+IntToStr(CountReg)+' отчетов', ICO_ERR)
                  else FFINFO.AddInfo('Импортировано '+IntToStr(CountReg)+' отчетов', ICO_INFO);
       If SMainReg<>'' then FFINFO.Caption := 'Импорт данных: '+SMainReg+' за '+SMonth+' '+SYear+' г.';
       FFINFO.Hide;
       FFINFO.AReport.Enabled := true;
       FFINFO.ShowModal;
       FFINFO.Free;
       ForegroundWindow;

       FFMAIN.Repaint;
    end;

    ModalResult := mrOk;
end;


{==============================================================================}
{============================   ACTION: ОТМЕНА   ==============================}
{==============================================================================}
procedure TFIMPORT_RAR.ACancelExecute(Sender: TObject);
begin
    ModalResult := mrCancel;
end;


{==============================================================================}
{=========================   ACTION: ВЫДЕЛИТЬ ВСЕ   ===========================}
{==============================================================================}
procedure TFIMPORT_RAR.ASelectAllExecute(Sender: TObject);
var I: Integer;
begin
    For I:=0 to CBView.Count - 1 do CBView.Checked[I]:=true;
    EnablAction;
end;


{==============================================================================}
{=========================   ACTION: ВЫДЕЛИТЬ ВСЕ   ===========================}
{==============================================================================}
procedure TFIMPORT_RAR.AUnselectAllExecute(Sender: TObject);
var I: Integer;
begin
    For I:=0 to CBView.Count - 1 do CBView.Checked[I]:=false;
    EnablAction;
end;

{==============================================================================}
{==========================   СОБЫТИЕ: ON_CLICK   =============================}
{==============================================================================}
procedure TFIMPORT_RAR.CBViewClick(Sender: TObject);
begin
    EnablAction;
end;

procedure TFIMPORT_RAR.CBViewKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    CBViewClick(Sender);
end;


{==============================================================================}
{==========================   СОБЫТИЕ: CB_CLICK   =============================}
{==============================================================================}
procedure TFIMPORT_RAR.CBSubRegClick(Sender: TObject);
begin
    WriteLocalBool(INI_SET, INI_SET_IMPORT_SUBREG, CBSubReg.Checked);
    EnablAction;
end;

procedure TFIMPORT_RAR.CBSubRegKeyPress(Sender: TObject; var Key: Char);
begin
    CBSubRegClick(Sender);
end;

procedure TFIMPORT_RAR.CBVerifyClick(Sender: TObject);
begin
    WriteLocalBool(INI_SET, INI_SET_IMPORT_VERIFY_RAR, CBVerify.Checked);
end;

procedure TFIMPORT_RAR.CBVerifyKeyPress(Sender: TObject; var Key: Char);
begin
    CBVerifyClick(Sender);
end;

end.
