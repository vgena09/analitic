unit MREPORT;

interface

uses
   Winapi.Windows,
   System.Classes, System.SysUtils, System.Variants, System.IniFiles,
   System.DateUtils, System.Math,
   Vcl.Dialogs, Vcl.Controls, IdGlobal, Vcl.ComCtrls, Vcl.OleServer,
   Data.DB, Data.Win.ADODB, Excel2000,
   MAIN, MPARAM, MBASE, FunType;

type
   TAnaliz = record
      PLParam               : PListParam;    // Указатель на список параметров произвольной аналитики, устанавливаемые при диалоге
   // LParam                : TListParam;    // Список параметров произвольной аналитики, устанавливаемых при диалоге
      LRegions              : TStringList;   // Список анализируемых регионов     При анализе региона: список доступных регионов - иеррархия: 1 - район, 2 - область, 3 - республика
      IStep                 : Integer;       // Шаг периода (в месяцах)    [-12]
      ILoop                 : Integer;       // Число периодов             [2]
      IRowBefore, IRowAfter : Integer;       // Число строк в шапке и подвале (только для переменного числа параметров)
      FirstColDat           : Integer;       // Первый столбец с данными      (только для переменного числа параметров)
      IPeriod               : Integer;       // Текущий период     ILoop * IStep
   end;
   PAnaliz = ^TAnaliz;

   TREPORT = class
   public
      constructor Create;
      destructor  Destroy; override;

      function  CreateReport(const SYear_, SMonth_, SRegion_, SFormula_ : String;
                             const IDReport_, IDMatrix_, CorrMonth_     : Integer;
                             const IsNew : Boolean): Boolean;
   private
      FFMAIN                           : TFMAIN;
      FPARAM                           : TFPARAM;
      FBASE                            : TBASE;
      WB                               : Variant;
      FIni                             : TMemIniFile;
      Analiz                           : TAnaliz;

      SYear, SMonth, SRegion, SFormula : String;
      IDReport, IDMatrix, CorrMonth    : Integer;

      SType, SPage, SBlock             : String;
      IPage                            : Integer;

      IsCtrl                           : Boolean;             // Нажата ли Crtl во время запуска функции

      function  CreatePage(const SPage_ : String): Boolean;
      function  CreateBlockStat: Boolean;
      function  CreateBlockAnal: Boolean;

      function  ValueCell(const SKey: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
      function  ValueCellStep(const SBlock_: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
      procedure SetInfo(const SDopText: String);
   end;

implementation

uses FunConst, FunExcel, FunIO, FunBD, FunSys, FunVcl, FunInfo, FunText, FunDay, FunSum;

{$INCLUDE MREPORT_STAT}
{$INCLUDE MREPORT_ANAL}


{==============================================================================}
{=============================   КОНСТРУКТОР   ================================}
{==============================================================================}
constructor TREPORT.Create;
begin
    inherited Create;

    {Инициализация}
    FFMAIN          := TFMAIN(GlFindComponent('FMAIN'));
    FPARAM          := TFPARAM(FFMAIN.FFPARAM);
    FBASE           := TBASE.Create;
    Analiz.LRegions := TStringList.Create;
    FFMAIN.Repaint;
end;


{==============================================================================}
{==============================   ДЕСТРУКТОР   ================================}
{==============================================================================}
destructor TREPORT.Destroy;
begin
    {Деинициализация}
    FBASE.Free;
    Analiz.LRegions.Clear;
    Analiz.LRegions.Free;
    WB := Unassigned;
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';

    inherited Destroy;
end;


{==============================================================================}
{=========                    СОЗДАТЬ ОТЧЕТ EXCEL                      ========}
{==============================================================================}
{=========  CorrMonth - коррекция месяца при создании нового отчета    ========}
{=========  Возврат   - допустимость контроля                          ========}
{==============================================================================}
function TREPORT.CreateReport(const SYear_, SMonth_, SRegion_, SFormula_ : String;
                              const IDReport_, IDMatrix_, CorrMonth_     : Integer;
                              const IsNew : Boolean): Boolean;
var LPag                                    : TStringList;
    IPage, I                                : Integer;
    SPathIni, SPathSrc, SPathDect, SPathDat : String;
begin
    {Инициализация}
    Result    := false;
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Формирование отчета: инициализация';
    SYear     := SYear_;
    SMonth    := SMonth_;
    SRegion   := SRegion_;
    SFormula  := SFormula_;
    IDReport  := IDReport_;
    IDMatrix  := IDMatrix_;
    CorrMonth := CorrMonth_;
    IsCtrl    := GetKeyState(VK_LCONTROL);
    If (Not IsIntegerStr(SYear)) or (MonthStrToInd(SMonth)<=0) then begin ErrMsg('Некорректный период: ['+SYear+' '+SMonth+']!'); Exit; end;
    If SRegion = ''                                            then begin ErrMsg('Не определен регион!'); Exit; end;
    If (IDMatrix < 0) or (IDMatrix > 2)                        then begin ErrMsg('Недопустимое значение IDMatrix!'); Exit; end;
    Analiz.PLParam := @FPARAM.LParam;

    {Файл-источник}
    SPathSrc := FFMAIN.FormsCounterToFile(IDReport);
    If Not FileExists(SPathSrc) then begin ErrMsg('Не найден шаблон отчета!'+CH_NEW+SPathSrc); Exit; end;

    {Файл индексов}
    SPathIni := ChangeFileExt(SPathSrc, '.ini');
    If Not FileExists(SPathIni) then begin ErrMsg('Не найден файл индексов!'+CH_NEW+SPathIni); Exit; end;

    {Файл-назначение}
    SPathDect := CopyFileToTemp(SPathSrc);
    If SPathDect='' then begin ErrMsg('Не определен файл отчета!'+CH_NEW+SPathDect); Exit; end;

    {Предупреждение о существовании старых таблиц отчета}
    If IsNew then begin
       SPathDat := FFMAIN.PathRegion(SYear, SMonth, SRegion);
       If FFMAIN.IsTableDatIncludeIni(SPathIni, SPathDat) then begin
          If MessageDlg('Банк данных содержит таблицы, которые при сохранении создаваемого отчета будут заменены на новые! Продолжить?', mtWarning, [mbYes, mbNo], 0)<>mrYes
          then Exit;
       end;
    end else CorrMonth := 0;

    {Инициализация INI}
    FIni := TMemIniFile.Create(SPathIni);
    LPag := TStringList.Create;
    try
       {Список листов}
       If Not SeparatMStr(FIni.ReadString(INI_PARAM, INI_PARAM_PAGES, ''), @LPag, ', ') then begin ErrMsg('Ошибка в списке листов отчета!'); Exit; end;

       {Предупреждение о большом отчете}
       If (Not IsNew) and (LPag.Count >= 50) then begin
          If MessageDlg('Отчет содержит '+ IntToStr(LPag.Count) +' обрабатываемых страниц и для его открытия может понадобиться время. Продолжить?', mtWarning, [mbYes, mbNo], 0)<>mrYes
          then Exit;
       end;

       try
          {Информатор}
          FFMAIN.PBar.Position := 0;
          FFMAIN.PBar.Visible  := true;
          FFMAIN.Repaint;

          {Инициализация EXCEL}
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' Формирование отчета: открытие EXCEL';
          FFMAIN.PBar.Position := 2;
          FFMAIN.WB.Disconnect;
          FFMAIN.XLS.Disconnect;
          FFMAIN.XLS.ConnectKind          := ckNewInstance;                     // ckRunningOrNew; - недопустимо т.к. проблемы
          FFMAIN.XLS.Connect;
          FFMAIN.XLS.AutoQuit             := false;
          FFMAIN.XLS.EnableEvents         := true;
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' Формирование отчета: открытие документа';
          FFMAIN.PBar.Position := 5;
          FFMAIN.XLS.Workbooks.Add(SPathDect, LOCALE_USER_DEFAULT);
          If FFMAIN.XLS.Workbooks.Count = 0 then begin ErrMsg('Ошибка открытия отчета!'); Exit; end;
          If (IDMatrix = 0) and (GetMSOfficeVer(msExcel) > 11) then FFMAIN.RunMacroSafe('ЭтаКнига.Workbook_Open'); // Для Office > 2003
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' Формирование отчета: настройка документа';
          FFMAIN.PBar.Position := 8;
          FFMAIN.XLS.Visible[LOCALE_USER_DEFAULT]             := false;         // Для Office > 2003
          Variant(FFMAIN.XLS.Application).UseSystemSeparators := false;         // Не работает для Office > 2003 (оставил на всякий случай)
          Variant(FFMAIN.XLS.Application).DecimalSeparator    := ',';           // ...
          Variant(FFMAIN.XLS.Application).ThousandsSeparator  := ' ';           // ...
          FFMAIN.XLS.UseSystemSeparators := false;                              // Исправляет ошибку на > 2003
          FFMAIN.XLS.DecimalSeparator    := ',';                                // ...
          FFMAIN.XLS.ThousandsSeparator  := ' ';                                // ...

          FFMAIN.WB.ConnectTo(FFMAIN.XLS.ActiveWorkbook);
          WB := FFMAIN.XLS.Workbooks[1];
          WB.Application.EnableEvents    := false;
          WB.Application.CellDragAndDrop := false;                              // Запретить перетаскивание ячеек

          {Записываем в отчет базовые год, месяц, регион}
          WB.WorkSheets[1].Unprotect;
          WB.WorkSheets[1].Range[CELL_YEAR].Value    := SYear;
          WB.WorkSheets[1].Range[CELL_MONTH].Value   := SMonth;
          WB.WorkSheets[1].Range[CELL_REGION].Value  := SRegion;
          WB.WorkSheets[1].Range[CELL_CAPTION].Value := FFMAIN.FormulaToCaption(SFormula, false);

          {Почередно формируем листы отчета}
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' Формирование отчета: заполнение документа';
          FFMAIN.PBar.Position := 10;
          For IPage:=0 to LPag.Count-1 do begin
             I := StrToInt(LPag[IPage]);
             If IsNew and (CorrMonth = 0) then begin
                {Создаем отчет}
                try If (IDMatrix = 0) and FIni.ReadBool(INI_PAGE+LPag[IPage], INI_PAGE_PROTECT, true) then begin
                       WB.WorkSheets[I].EnableSelection := xlUnlockedCells;
                       WB.WorkSheets[I].Protect('');
                    end;
                except end;
             end else begin
                {Открываем отчет}
                If LPag.Count > 1 then FFMAIN.PBar.Position := 10 + (85 * IPage Div (LPag.Count-1))
                                  else FFMAIN.PBar.Position := 55;
                If Not CreatePage(LPag[IPage]) then begin ErrMsg('Ошибка обработки листа отчета № '+LPag[IPage]+'!'); Exit; end;
             end;
             {Для Office > 2003 корректировка}
             If GetMSOfficeVer(msExcel) > 11 then begin
                WB.WorkSheets[I].EnableFormatConditionsCalculation := true;     // Разрешить обработку условного форматирования
                WB.WorkSheets[I].EnableCalculation := true;                     // Разрешить пересчет
             end;
          end;

          {Информатор}
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Формирование отчета: завершение операции';
          FFMAIN.PBar.Position := 100;

          {Ярлычки листов}
          If IDMatrix>0 then WB.Application.ActiveWindow.DisplayWorkbookTabs := true
                        else WB.Application.ActiveWindow.DisplayWorkbookTabs := FIni.ReadBool(INI_PARAM, INI_PARAM_PAGELINKS, false);

          {Разрешаем обработку событий в EXCEL}
          WB.Application.EnableEvents:=true;

          {Корректируем область просмотра, раскраску}
          If (IDMatrix = 0) then FFMAIN.RunMacroSafe('RunAfterCreate');

          {Для матрицы отключаем все виды контроля}
          If (IDMatrix > 0) then begin
             FFMAIN.XLS.DisplayFullScreen[LOCALE_USER_DEFAULT]:=false;
             FFMAIN.XLS.WindowState[0] := xlMaximized;
          end;

          {Восстановление активного листа и ячейки}
          LoadActiveCell(FFMAIN.XLS.Workbooks[1]);

          {Настройки EXCEL}
          FFMAIN.XLS.Visible[LOCALE_USER_DEFAULT] := true;
          If GetMSOfficeVer(msExcel)>=11 then SetForegroundWindow(Variant(FFMAIN.XLS.DefaultInterface).Hwnd);
          FFMAIN.XLS.Workbooks[1].Saved[LOCALE_USER_DEFAULT]:=true;
          FFMAIN.XLS.UserControl := true;

          {Возвращаемый результат}
          Result:=true;
       finally
          If Result then begin
             Result:=IS_ADMIN and (IDMatrix=0) and (Not FFMAIN.IsFormReadOnly(IDReport_));
          end else begin
             try FFMAIN.WB.Close(false); except end;
             FFMAIN.WB.Disconnect;
             FFMAIN.XLS.Disconnect;
          end;

          FFMAIN.PBar.Visible  := false;
          FFMAIN.PBar.Position := 0;
       end;

   finally
       LPag.Free;
       FIni.Free;
       WB     := Unassigned;
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
   end;
end;


{==============================================================================}
{=======================  СОЗДАТЬ СТРАНИЦУ ОТЧЕТА EXCEL  ======================}
{==============================================================================}
function TREPORT.CreatePage(const SPage_ : String): Boolean;
var SBlockList : String;
begin
    {Инициализация}
    Result := false;
    SPage  := SPage_;
    If Not IsIntegerStr(SPage) then begin ErrMsg('Недопустимый номер листа: '+SPage+'!'); Exit; end;
    IPage  := StrToInt(SPage);
    FFMAIN.Repaint;

    {Допустимость номера листа}
    If IPage > Integer(WB.WorkSheets.Count) then begin ErrMsg('Недопустимый номер листа: '+SPage+'!'); Exit; end;

    {Тип отчета}
    SType := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_TYPE, '');

    {Список обрабатываемых блоков текущего листа}
    Case IDMatrix of
    0, 2 : SBlockList := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_OUT, '');
    1    : SBlockList := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_IN, '');
    else begin ErrMsg('Недопустимое значение IDMatrix!'); Exit; end;
    end;

    {Обрабатываем блоки текущего листа}
    WB.WorkSheets[IPage].Unprotect;
    try
       SBlock := CutBlock(SBlocklist, ',');
       While SBlock<>'' do begin
          If SType = '' then begin If Not CreateBlockStat then Exit; end
                        else begin If Not CreateBlockAnal then Exit; end;
          SBlock := CutBlock(SBlockList, ',');
       end;
    finally
       If (IDMatrix = 0) and FIni.ReadBool(INI_PAGE+SPage, INI_PAGE_PROTECT, true) then begin
          WB.WorkSheets[IPage].EnableSelection := xlUnlockedCells;
          WB.WorkSheets[IPage].Protect('');
       end;
    end;

    {Возвращаемый результат}
    Result:=true;
end;



{******************************************************************************}
{**************************  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ  *************************}
{******************************************************************************}


{==============================================================================}
{==========================  ОПРЕДЕЛЯЕТ ЗНАЧЕНИЕ ЯЧЕЙКИ   =====================}
{==============================================================================}
function TREPORT.ValueCell(const SKey: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
var SKey0  : String;
    SBlock : String;
    SVal   : String;
begin
    {Инициализация}
    Result := 0;
    SKey0  := SKey;

    //{При необходимости корректируем период и регион}
    //If CorrPeriod <> 0  then SKey0 := SKey0 + '/' + IntToStr(CorrPeriod);
    //If CorrRegion <> '' then SKey0 := SKey0 + '/' + CorrRegion;

    try
       {При наличии хоть одного блока - алгоритм полной обработки}
       SBlock := CutModulChar(SKey0, '[', ']');
       If SBlock<>'' then begin
          While SBlock<>'' do begin
             SVal   := FloatToStr(ValueCellStep(SBlock, CorrPeriod, CorrRegion));
             SKey0  := ReplModulChar(SKey0, SVal, '[', ']');
             SBlock := CutModulChar(SKey0,        '[', ']');
          end;

          {Вычисляем значение}
          Result:=CalcStr(SKey0);

       {При отсутствии блоков - упрощенный алгоритм}
       end else begin
          Result := ValueCellStep(SKey, CorrPeriod, CorrRegion);
       end;
    finally
    end;
end;


{==============================================================================}
{==========================  ОПРЕДЕЛЯЕТ ЗНАЧЕНИЕ ЯЧЕЙКИ   =====================}
{==============================================================================}
function TREPORT.ValueCellStep(const SBlock_: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
var IDTab, IDRow, IDCol, IDReg, IDPer : String;
    SBlock                            : String;
    SPath                             : String;
begin
    {Инициализация}
    Result := 0;
    SBlock := SBlock_;
    try
       IDTab  := CutBlock(SBlock, CH_SPR);
       If CmpStr(IDTab, 'КОД') then begin
          SBlock := SFormula+CH_SPR+SBlock;
          IDTab  := CutBlock(SBlock, CH_SPR);
       end;
       IDCol  := CutBlock(SBlock, CH_SPR);
       IDRow  := CutBlock(SBlock, CH_SPR);
       IDPer  := CutBlock(SBlock, CH_SPR); If IDPer='' then IDPer:='0';
       IDReg  := CutBlock(SBlock, CH_SPR);
       SBlock := IDTab+CH_SPR+IDCol+CH_SPR+IDRow;
       If CorrPeriod <> 0 then SBlock := SBlock+CH_SPR+IntToStr(StrToInt(IDPer)+CorrPeriod)
                          else SBlock := SBlock+CH_SPR+IDPer;
       If CorrRegion <>'' then SBlock := SBlock+CH_SPR+CorrRegion
                          else SBlock := SBlock+CH_SPR+IDReg;
       // Задача на перспективу: в целях оптимизации коррекцию региона произвести сейчас
       SPath  := FFMAIN.PathDat(SYear, SMonth, SRegion, SBlock);
       Result := FBASE.GetVal(SPath, IDTab, IDCol, IDRow);
    finally end;
end;


{==============================================================================}
{================================  ИНФОРМАТОР  ================================}
{==============================================================================}
procedure TREPORT.SetInfo(const SDopText: String);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Формирование отчета: [ '+SRegion+' ] '+SPage+'!'+SBlock+' '+SDopText;
    FFMAIN.StatusBar.Repaint;
end;


end.
