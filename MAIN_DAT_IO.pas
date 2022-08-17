{==============================================================================}
{===========    !!! ВАЖНО !!! ВАЖНО !!! ВАЖНО !!! ВАЖНО !!!     ===============}
{==============================================================================}
{===========  РАБОТА С АБСОЛЮТНЫМИ ПУТЯМИ, YMR НЕ ИСПОЛЬЗУЕТСЯ  ===============}
{==============================================================================}

{$INCLUDE MAIN_DAT_IO_SECTION}

{==============================================================================}
{=============   ЧИТАЕТ ЗНАЧЕНИЕ ИЗ DAT-ФАЙЛА БЕЗ КЭШИРОВАНИЯ   ===============}
{==============================================================================}
function TFMAIN.GetValNoCash(const SPathDat, STab, SCol, SRow: String): String;
var FDat: TIniFile;
begin
    {Инициализация}
    Result:='';
    If Not FileExists(SPathDat) then Exit;

    {Список таблиц}
    FDat:=TIniFile.Create(SPathDat);
    try      Result:=FDat.ReadString(STab, SCol+' '+SRow, '');
    finally FDat.Free;
    end;
end;


{==============================================================================}
{====================   СУЩЕСТВУЕТ ЛИ СЕКЦИЯ DAT-ФАЙЛА   ======================}
{==============================================================================}
function TFMAIN.IsExistsTab(const SPath, STab: String): Boolean;
var LTab : TStringList;
begin
    {Инициализация}
    Result := false;

    {Проверка}
    LTab := TStringList.Create;
    try     If Not GetListTabDat(@LTab, SPath) then Exit;
            Result := (LTab.IndexOf(STab) >= 0);
    finally LTab.Free;
    end;
    // Result:=FDat.SectionExists(STab);   // не определяет пустую секцию
end;


{==============================================================================}
{===================   ИМЕЕТСЯ ЛИ ЗАДАННОЕ ЗНАЧЕНИЕ СТРОКИ  ===================}
{==============================================================================}
{===  SRow = '' - любое значение (есть ли какое-либо значение для таблицы) ====}
{==============================================================================}
function TFMAIN.IsExistsRow(const SPath, STab, SRow: String): Boolean;
var SList : TStringList;
    I     : Integer;
begin
    {Инициализация}
    Result := false;
    If Not FileExists(SPath) then Exit;

    SList := TStringList.Create;
    try
        {Читаем секцию}
        If Not GetTabVal(SPath, STab, @SList) then Exit;
        {Просматриваем секцию}
        If SRow='' then begin
           Result := (SList.Count > 0);
        end else begin
           For I:=0 to SList.Count-1 do begin
              If CutSlovo(SList[I], 2, ' ') = SRow then begin
                 Result:=true;
                 Break;
              end;
           end;
        end;
    finally
       SList.Free;
    end;
end;



{==============================================================================}
{===  НАЛИЧИЕ В DAT-ФАЙЛЕ ХОТЬ ОДНОЙ ТАБЛИЦЫ, ЗАДЕКЛАРИРОВАННОЙ В INI-ФАЙЛЕ ===}
{==============================================================================}
function TFMAIN.IsTableDatIncludeIni(const SPathIni, SPathDat: String): Boolean;
var LIni, LDat : TStringList;
    I          : Integer;
begin
    {Инициализация}
    Result:=false;
    If Not FileExists(SPathDat) then Exit;
    LIni := TStringList.Create;
    LDat := TStringList.Create;
    try
       {Списки таблиц}
       If Not GetListTabIni(@LIni, SPathIni) then begin ErrMsg('Ошибка чтения списка таблиц из INI-файла!'+CH_NEW+SPathIni); Exit; end;
       If Not GetListTabDat(@LDat, SPathDat) then begin ErrMsg('Ошибка чтения списка таблиц из DAT-файла!'+CH_NEW+SPathDat); Exit; end;

       {Проверка}
       For I:=0 to LDat.Count-1 do begin
          If LIni.IndexOf(LDat[I])>=0 then begin
             Result:=true;
             Break;
          end;
       end;
    finally
       LDat.Free;
       LIni.Free;
    end;
end;


{==============================================================================}
{============               ТОЛЬКО ЛИ ФОРМА ДЛЯ ЧТЕНИЯ             ============}
{============      (ПО ПРЕФИКСУ И НАЛИЧИЮ БЛОКА ВВОДА INI ФАЙА)    ============}
{==============================================================================}
function TFMAIN.IsFormReadOnly(const IDForm: Integer): Boolean;
var FIni                : TMemIniFile;
    LPage               : TStringList;
    SPathForm, SPathIni : String;
    S                   : String;
    I                   : Integer;
begin
    {Инициализация}
    Result := (FormsCounterToBankPrefix(IDForm) <> '');
    If Result then Exit else Result := true;

    SPathForm := FormsCounterToFile(IDForm);
    If Not FileExists(SPathForm) then Exit;
    SPathIni := ChangeFileExt(SPathForm, '.ini');
    If Not FileExists(SPathIni) then Exit;

    FIni  := TMemIniFile.Create(SPathIni);
    LPage := TStringList.Create;
    try
       {Читаем список листов}
       S:=FIni.ReadString(INI_PARAM, INI_PARAM_PAGES, '');
       If Not SeparatMStr(S, @LPage, ', ') then Exit;

       {Просматриваем каждый лист на возможность ввода данных}
       For I:=0 to LPage.Count-1 do begin
          S:=FIni.ReadString(INI_PAGE+LPage[I], INI_PAGE_IN, '');
          If S<>'' then begin
             Result:=false;
             Break;
          end;
       end;
    finally
       LPage.Free;
       FIni.Free;
    end;
end;


