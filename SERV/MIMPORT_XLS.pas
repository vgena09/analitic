unit MIMPORT_XLS;

interface

uses
   Winapi.Windows {Beep},
   System.Classes, System.SysUtils, System.Variants, System.IniFiles,
   Vcl.Dialogs, Vcl.Controls, IdGlobal,
   Data.DB, Data.Win.ADODB,
   MAIN, FunType;

type
   TIMPORT_XLS = class
   public
      constructor Create;
      destructor  Destroy; override;

      function  ImportXls(const SPath: String): Boolean;
      function  ImportXlsID(const WB: Variant; const IsDialog: Boolean): Boolean;
   private
      FFMAIN  : TFMAIN;
   end;

implementation

uses FunConst, FunExcel, FunIO, FunBD, FunSys, FunVcl, FunText, FunDay, FunSum,
     FunVerify, FunIni,
     MTITLE, MBASE;


{==============================================================================}
{=============================   КОНСТРУКТОР   ================================}
{==============================================================================}
constructor TIMPORT_XLS.Create;
begin
    inherited Create;

    FFMAIN  := TFMAIN(GlFindComponent('FMAIN'));
end;


{==============================================================================}
{==============================   ДЕСТРУКТОР   ================================}
{==============================================================================}
destructor TIMPORT_XLS.Destroy;
begin
    // ...

    inherited Destroy;
end;


{==============================================================================}
{=====================  НАДСТРОЙКА: ИМПОРТ ФОРМЫ XLS  =========================}
{==============================================================================}
function TIMPORT_XLS.ImportXls(const SPath: String): Boolean;
var E: Variant;
begin
    {Инициализация}
    Result:=false;

    {Открываем Excel}
    If Not CreateExcel then begin ErrMsg('Ошибка открытия MS Excel!'); Exit; end;
    try

       {Открываем файл с блокировкой макросов}
       If Not OpenXls(SPath, true) then begin ErrMsg('Ошибка открытия файла: '+ExtractFileName(SPath)); Exit; end;
       try

          {Импорт отчета}
          E      := GetExcelApplicationID;
          E.Workbooks[1].Windows[1].Visible:=true;     //При открытии скрытой книги LOCALE_USER_DEFAULT
          Result := ImportXlsID(E.ActiveWorkbook, true);

       {Закрываем файл}
       finally
          E.DisplayAlerts:=false;
          If Not CloseXls then ErrMsg('Ошибка закрытия импортируемого отчета!');
       end;

    {Закрываем Excel}
    finally
       If Not CloseExcel then ErrMsg('Ошибка закрытия Excel!');
       E  := Unassigned;
    end;
end;


{==============================================================================}
{============================  ИМПОРТ ФОРМЫ XLS  ==============================}
{==============================================================================}
function TIMPORT_XLS.ImportXlsID(const WB: Variant; const IsDialog: Boolean): Boolean;
label Dl, Er;
var FTITLE                        : TFTITLE;
    FBASE                         : TBASE;
    LForma_                       : TLForma;
    FIni                          : TMemIniFile;
    LIni                          : TStringlist;
    FDate                         : TDate;
    SPathIni, SPathDat            : String;
    S, SLog                       : String;
    SForm, SYear, SMonth, SRegion : String;
    SPageList,  SPage             : String;
    SBlockList, SBlock            : String;
    I, IPage                      : Integer;
    IsVerify                      : Boolean;


    {==========================================================================}
    {==================   Осуществляет ввод области    ========================}
    {==========================================================================}
    {==================   SBlock = 'A5:BH125'          ========================}
    {==================   SBlock = S1+':'+S2           ========================}
    {==================   SBlock = Row1,Col1,Row2,Col2 ========================}
    {==========================================================================}
    function InputBlock(const IPage: Integer; const SBlock: String): Boolean;
    var S, S1, S2              : String;
        SVal                   : Variant;
        IDTab, IDRow, IDCol    : String;
        Row1, Col1, Row2, Col2 : Integer;
        IRow, ICol             : Integer;
        AData                  : Variant;
    begin
        {Инициализация}
        Result:=false;
        S1 := AnsiUpperCase(CutSlovo(SBlock, 1, ':'));
        S2 := AnsiUpperCase(CutSlovo(SBlock, 2, ':'));
        If S1='' then Exit;
        If S2='' then S2:=S1;

        {Информатор}
        FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Импорт: [ '+SRegion+' ] '+IntToStr(IPage)+'!'+SBlock;

        {Определяем область}
        If Not ExcelSeparatAddress(S1, Col1, Row1) then Exit;
        If Not ExcelSeparatAddress(S2, Col2, Row2) then Exit;
        If (Col1>Col2) or (Row1>Row2) then Exit;
        AData := VarArrayCreate([Row1, Row2, Col1, Col2], varVariant);
        AData := WB.WorkSheets[IPage].Range[S1, S2].Value;      // Некорректная ф-ция: при S1=S2 возвращает не массив а значение, в т.ч. и Unassigned; изменяет начало/конец массива: AData[IRow, ICol] - не верно

        try
           {Просматриваем все ячейки области}
           For IRow := Row1 to Row2 do begin
              For ICol := Col1 to Col2 do begin
                 {Переводим адрес ячейки в форму A1}
                 S:=ExcelCombineAddress(ICol, IRow);
                 If S='' then Exit;
                 {Читаем код для текущей ячейки: 5/4/21/0/0}
                 S:=FIni.ReadString(INI_PAGE+SPage, S, '');
                 If S='' then Continue;
                 {Разбиваем код на составляющие}
                 IDTab := CutSlovo(S, 1, CH_SPR);
                 IDCol := CutSlovo(S, 2, CH_SPR);
                 IDRow := CutSlovo(S, 3, CH_SPR);
                 {Записываем значение в DAT-файл из области}
                 If S1<>S2 then begin
                    SVal := AData[IRow-Row1+1, ICol-Col1+1];       // AData[IRow, ICol] - не верно
                 end else begin
                    If Not VarIsEmpty(AData) then SVal := AData else SVal := '';
                 end;
                 If VarType(SVal)=varError then SVal:='';
                 If Not FBASE.SetVal(SPathDat, IDTab, IDCol, IDRow, SVal) then Exit;
              end;
           end;
           Result:=true;
        finally
           {Освобождаем память}
           VarClear(AData);
        end;
    end;


    {==========================================================================}
    {========       РАЗБИВАЕТ СТРОКУ-ПЕРИОД НА СОСТАВЛЯЮЩИЕ        ============}
    {==========================================================================}
    {========  SPeriod - '(за 9 месяцев 2008 г. в сравнении ...'   ============}
    {========  SPeriod - 'за 2 месяца 2008 г.' 'за 2 месяца 2008г.'============}
    {========  SPeriod - 'за полугодие 2008 г.'                    ============}
    {========  SPeriod - 'за 2008 год'                             ============}
    {==========================================================================}
    function SeparatOldPeriod(const SPeriod: String; var SYear, SMonth: String): Boolean;
    var S, S1 : String;
        I     : Word;
    begin
        {Инициализация}
        Result := false;
        SYear  := '';
        SMonth := '';
        S      := Trim(SPeriod);

        {S = '(за 9 месяцев 2008' или 'за 2008'}
        I := Pos(' г.' , S); If I>4 then Delete(S, I, Length(S));
        I := Pos(' год', S); If I>4 then Delete(S, I, Length(S));
        I := Pos('г.' , S);  If I>3 then Delete(S, I, Length(S));
        I := Pos('год', S);  If I>3 then If Pos('полугодие', S)=0 then Delete(S, I, Length(S));
        S := Trim(S);

        {Вырезаем год}
        S1 := TokCharEnd(S, ' ');
        If (Not IsIntegerStr(S1)) or (Length(S1)<>4) then Exit;
        SYear := S1;

        {Вырезаем месяц}
        {S = '(за 9 месяцев' или 'за'}
        TokChar(S, ' ');
        If S<>'' then begin
           I := MonthOldStrToInd(S);   If (I<Low(MonthList)) or (I>High(MonthList)) then Exit;
           SMonth := MonthList[I];
        end else begin
           SMonth := MonthList[High(MonthList)];
        end;

        {Возвращаемый результат}
        Result:=true;
    end;

begin
    {Инициализация}
    Result := false;

    {Получаем тип формы}
    SLog := WB.WorkSheets[1].Range[CELL_LOG].Value;

    {**************************************************************************}
    {******************   ПРОКУРОРСКАЯ ОТЧЕТНОСТЬ   ***************************}
    {**************************************************************************}
    If CmpStr(IDLOG,  Copy(SLog, 1, Length(IDLOG))) then begin
       {Корректируем идентификатор}
       Delete(SLog, 1, Length(IDLOG));
       SLog:=Trim(SLog);

       {Определяем базовые год, месяц, регион}
       SYear   := WB.WorkSheets[1].Range[CELL_YEAR].Value;
       SMonth  := WB.WorkSheets[1].Range[CELL_MONTH].Value;
       SRegion := WB.WorkSheets[1].Range[CELL_REGION].Value;
       Goto Dl;
    end;

    {**************************************************************************}
    {**************************   ОТЧЕТНОСТЬ МВД  *****************************}
    {**************************************************************************}
    // If WB.WorkSheets.Count<5 then Goto Er;  701 форма содержит 4 листа
    SLog := '';
    try
       {Выбираем все формы с идентификаторами}
       If FFMAIN.FindForms(@LForma_, false, true, -1, -1, '', '', '', '', true, 0) then begin
          For I:=High(LForma_) downto Low(LForma_) do begin
             {Сравниваем идентификатор}
             SForm := ReadWBCells(WB, LForma_[I].FCellID);
             If Not CmpStr(SForm, LForma_[I].FStrID) then Continue;

             {Читаем период и регион}
             S       := ReadWBCells(WB, LForma_[I].FCellPeriod);
             SRegion := ReadWBCells(WB, LForma_[I].FCellRegion);

             {Ищем индекс формы в зависимости от периода}
             SeparatOldPeriod(S, SYear, SMonth);
             FDate := SDatePeriod(SYear, SMonth);
             If LForma_[I].FBegin > 0 then If LForma_[I].FBegin > FDate then Continue;
             If LForma_[I].FEnd   > 0 then If LForma_[I].FEnd   < FDate then Continue;

             {Индекс формы найден}
             SLog := IntToStr(LForma_[I].FCounter);
             Break;
          end;
       end;
    finally
       SetLength(LForma_, 0);
    end;
    If SLog <> '' then Goto Dl;

    {Отчет не идентифицирован}
Er: ErrMsg('Файл не опознан как источник данных!');
    Exit;

Dl: {Допустимость формы}
    If (Pos(' ', SLog)>0) or (Not IsIntegerStr(SLog)) then begin ErrMsg('Тип формы не определен: '+SLog); Exit; end;

    {Определяем ini-файл формы - SPathIni}
    SForm    := FFMAIN.FormsCounterToCaption(StrToInt(SLog));
    If SForm='' then begin ErrMsg('Тип формы не определен: '+SLog); Exit; end;
    S        := FFMAIN.FormsCounterToFile(StrToInt(SLog));
    SPathIni := ChangeFileExt(S, '.ini');

    {Открываем ini-файл}
    If Not FileExists(SPathIni) then begin ErrMsg('Не найден ini-файл: '+SPathIni); Exit; end;
    FIni  := TMemIniFile.Create(SPathIni);
    FBASE := TBASE.Create;
    LIni  := TStringList.Create;
    try
       {Читаем список задействованных таблиц}
       If Not FFMAIN.GetListTabIni(@LIni, SPathIni) then begin ErrMsg('Ошибка чтения списка таблиц из INI-файла!'+CH_NEW+SPathIni); Exit; end;

       {Подверждение-корректировка}
       FTITLE:=TFTITLE.Create(nil);
       try     If Not FTITLE.Execute(SYear, SMonth, SRegion, SForm, false, IsDialog) then Exit;
       finally FTITLE.Free;
       end;
       FFMAIN.Repaint;

       {Предупреждение о существовании старых таблиц отчета}
       SPathDat:=FFMAIN.PathRegion(SYear, SMonth, SRegion);
       //YMR := FFMAIN.SetYMR(SYear, SMonth, SRegion, '');
       //If FFMAIN.IsTableDatIncludeIni(SPathIni, SPathDat) then begin
       //   If MessageDlg('Банк данных содержит таблицы, которые при импорте отчета будут заменены на новые! Продолжить?', mtWarning, [mbYes, mbNo], 0)<>mrYes
       //   then Exit;
       //end;

       {Создаем иерархию dat-файлов}
       S := FFMAIN.CreateRegions(SYear, SMonth, SRegion);
       If S='' then Exit;

       {Корректируем выделение}
       WriteLocalString(INI_SELECT, INI_SELECT_PATH, S);

       {Читаем признак проверки XLS-отчетов}
       IsVerify := ReadLocalBool(INI_SET, INI_SET_IMPORT_VERIFY_XLS, true);

       {Читаем список листов}
       SPageList:=FIni.ReadString(INI_PARAM, INI_PARAM_PAGES, '');

       {Просматриваем указанные листы}
       SPage:=CutBlock(SPageList, ',');
       While SPage<>'' do begin
          {Допустимость номера листа}
          If Not IsIntegerStr(SPage) then begin ErrMsg('Неверный номер листа ['+SPage+'] !'); Exit; end;
          IPage := StrToInt(SPage);

          {Перерисовка главного окна при длительной операции}
          FFMAIN.Repaint;

          {Читаем область ввода}
          SBlockList:=FIni.ReadString(INI_PAGE+SPage, INI_PAGE_IN, '');

          {ОТЛАДКА: Добавление строки}
          // WB.WorkSheets[IPage].Range['A2','A2'].Insert(3);

          {Просматриваем все области ввода текущего листа}
          SBlock:=CutBlock(SBlocklist, ',');
          While SBlock<>'' do begin
             If Not InputBlock(IPage, SBlock) then begin ErrMsg('Ошибка чтения блока ['+SPage+'!'+SBlock+'] !'); Exit; end;
             SBlock:=CutBlock(SBlockList, ',');
          end;

          {Следующий лист}
          SPage:=CutBlock(SPageList, ',');
       end;

       {Сохраняем таблицы}
       FBASE.SaveMemFile;
       FBASE.ClearCash;

       {Пересчитываем таблицы с рекурсией}
       For I:=0 to LIni.Count-1 do SumTable(LIni[I], SYear, SMonth, SRegion);

       {Проверяем таблицы}
       If IsVerify then VerifyTable(@LIni, SYear, SMonth, SRegion, true);
    finally
       LIni.Free;
       FBASE.Free;
       FIni.Free;
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
    end;

    {Возвращаемый результат}
    Result:=true;
end;

end.
