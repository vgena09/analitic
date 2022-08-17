{==============================================================================}
{==========================   ЧИТАЕМ ЗНАЧЕНИЕ   ===============================}
{==============================================================================}
function TBASE.GetValYMR(const YMR, STab, SCol, SRow: String): Extended;
var SPath : String;
begin
    SPath  := FFMAIN.TabToDatYMR(STab, YMR);
    Result := GetVal(SPath, STab, SCol, SRow);
end;


{==============================================================================}
{==========================   ЧИТАЕМ ЗНАЧЕНИЕ   ===============================}
{==============================================================================}
function TBASE.GetVal(const SPath, STab, SCol, SRow: String): Extended;
var FIni                            : TMemIniFile;
    SPath0                          : String;
    SKey, SKeyTab, SKeyCol, SKeyRow : String;
    LORKey_                         : TLORKey;
    S                               : String;
    I                               : Integer;

    SYear0, SMonth0, SRegion0       : String;
    SYear,  SMonth                  : String;
    DMonth                          : Integer;
begin
    {Инициализация}
    Result := 0;
    SPath0 := SPath;               // D:\База данных\2009\июнь\Берестовицкий.dat

    {*** Если значение является социально-экономическим показателем ***********}
    If FFMAIN.LCEP.IndexOf(STab) >= 0 then begin
       {*** Оптимизация включена (медленная работа) ***************************}
       If BSEPDetail then begin
          {Инициализация}
          SKey     := SCol+' '+SRow;
          SRegion0 := TokCharEnd(SPath0, '\');
          SMonth0  := TokCharEnd(SPath0, '\');
          SYear0   := TokCharEnd(SPath0, '\');
          {Просматриваем периоды}
          For DMonth:=0 to ISEPMonth do begin
             {Корректируем период}
             SYear  := SYear0;
             SMonth := SMonth0;
             If Not CorrectPeriodStr(SYear, SMonth, DMonth*(-1)) then Break;
             {Читаем значение}
             FIni:=TMemIniFile.Create(SPath0+'\'+SYear+'\'+SMonth+'\'+SRegion0);
             try     Result:=FIni.ReadInteger(STab, SKey, 0);
             finally FIni.Free;
             end;
             {Проверяем значение}
             If Result <> 0 then Break;
          end;
          {Выход}
          Exit;
       {*** Оптимизация выключена *********************************************}
       end else begin
          {Попытка чтения значения за январь текущего года}
          SPath0 := ReplSlovoEndChar(SPath0, 'январь', 2, '\');
          GetCash(SPath0, STab, SCol, SRow, Result);
          If Result <> 0 then Exit;
          {Попытка чтения значения за январь предыдущего года}
          SYear0 := CutSlovoEndChar(SPath0, 3, '\');
          try SYear0 := IntToStr(StrToInt(SYear0)-1); finally end;
          SPath0 := ReplSlovoEndChar(SPath0, SYear0, 3, '\');
       end;
    end;

    {*** Чтение данных из кэша ************************************************}
    If GetCash(SPath0, STab, SCol, SRow, Result) then If Result <> 0 then Exit;                       // If Result <> 0 может замедлить работу

    {*** Альтернативное чтение данных *****************************************}
    S := STab+' '+SCol+' '+SRow;
    try
       {Просмотр альтернатив}
       If FFMAIN.FindORKeys(@LORKey_, false, S, '') then begin
          For I:=Low(LORKey_) to High(LORKey_) do begin
             SKey    := LORKey_[I].FKey2;
             SKeyTab := TokChar(SKey, ' ');
             SKeyCol := TokChar(SKey, ' ');
             SKeyRow := TokChar(SKey, ' ');
             If GetCash(SPath, SKeyTab, SKeyCol, SKeyRow, Result) then If Result <> 0 then Exit;     // If Result <> 0 может замедлить работу
          end;
       end;
       If FFMAIN.FindORKeys(@LORKey_, false, '', S) then begin
          For I:=Low(LORKey_) to High(LORKey_) do begin
             SKey    := LORKey_[I].FKey1;
             SKeyTab := TokChar(SKey, ' ');
             SKeyCol := TokChar(SKey, ' ');
             SKeyRow := TokChar(SKey, ' ');
             If GetCash(SPath, SKeyTab, SKeyCol, SKeyRow, Result) then If Result <> 0 then Exit;    // If Result <> 0 может замедлить работу
          end;
       end;
    finally
       SetLength(LORKey_, 0);
    end;
end;


{==============================================================================}
{===================       ПОИСК DAT-ФАЙЛА В КЭШЕ       =======================}
{==============================================================================}
{===================   SPath  - путь к искомому файлу   =======================}
{===================   RESULT - индекс кэша             =======================}
{==============================================================================}
function TBASE.FindCashDat(const SPath: String): Integer;
var I      : Integer;
    SPath0 : String;
begin
    Result := -1;
    SPath0 := AnsiUpperCase(SPath);
    For I:=Low(CashTab) to High(CashTab) do begin
       If AnsiUpperCase(CashTab[I].SPath) = SPath0 then begin
          Result := I;
          Break;
       end;
    end;
end;


{==============================================================================}
{===================       ПОИСК ТАБЛИЦЫ В КЭШЕ         =======================}
{==============================================================================}
{===================   STab   - искомая таблица         =======================}
{===================   RESULT - индекс кэша             =======================}
{==============================================================================}
function TBASE.FindCashTab(const SPath, STab: String): Integer;
var I      : Integer;
    SPath0 : String;
begin
    Result := -1;
    SPath0 := AnsiUpperCase(SPath);
    For I:=Low(CashTab) to High(CashTab) do begin
       If (AnsiUpperCase(CashTab[I].SPath) = SPath0) and (CashTab[I].STab = STab) then begin
          Result := I;
          Break;
       end;
    end;
end;


{==============================================================================}
{===================   ЗАГРУЗКА DAT-ФАЙЛА В КЭШ         =======================}
{==============================================================================}
function TBASE.LoadCashDat(const SPath: String): Boolean;
var LMem           : TStringList;
    I, IMem        : Integer;
    SVal, STabCash : String;
begin
    {Инициализация}
    Result := false;

    {Ищем ранее кэшированный dat-файл}
    Result := (FindCashDat(SPath) >=0 );
    If Result then Exit;

    {При отсутствии кэширования добавляем dat-файл полностью}
    If Not FileExists(SPath) then Exit;
    I    := 0;
    LMem := TStringList.Create;
    try
       LMem.LoadFromFile(SPath);
       For IMem:=0 to LMem.Count-1 do begin
          SVal := LMem[IMem];
          STabCash := CutModulChar(SVal, '[', ']');
          If STabCash <> '' then begin
             {Кэшируем очередную таблицу}
             I := Length(CashTab);
             SetLength(CashTab, I+1);
             CashTab[I].SPath := SPath;
             CashTab[I].STab  := STabCash;
             CashTab[I].SList := TStringList.Create;
             Result := true;
          end else begin
             {Кэшируем значение таблицы}
             CashTab[I].SList.Add(SVal);
          end;
       end;
    finally
       LMem.Free;
    end;
end;


{==============================================================================}
{========================   ЧТЕНИЕ ЗНАЧЕНИЯ ИЗ КЭША   =========================}
{==============================================================================}
function TBASE.GetCash(const SPath, STab, SCol, SRow: String; var IVal: Extended): Boolean;
var S : String;
    I : Integer;
begin
    {Инициализация}
    Result := false;
    IVal   := 0;

    {Ищем кэш таблицы}
    I := FindCashTab(SPath, STab);

    {При ее отсутствии в кэше пытаемся кэшировать весь DAT-файл}
    If I < 0 then begin
       If Not LoadCashDat(SPath) then Exit;
       {После кеширования повторно ищем STab}
       I := FindCashTab(SPath, STab);
       If I < 0 then Exit;
    end;

    {Читаем значение}
    S := CashTab[I].SList.Values[SCol+' '+SRow];
    If IsFloatStr(S) then IVal := StrToFloat(S);

    {Возвращаемый результат}
    Result := true;
end;


