{==============================================================================}
{===========             РАБОТА С СЕКЦИЯМИ (ТАБЛИЦАМИ)          ===============}
{==============================================================================}
{===========    !!! ВАЖНО !!! ВАЖНО !!! ВАЖНО !!! ВАЖНО !!!     ===============}
{==============================================================================}
{===========  РАБОТА С АБСОЛЮТНЫМИ ПУТЯМИ, YMR НЕ ИСПОЛЬЗУЕТСЯ  ===============}
{==============================================================================}


{==============================================================================}
{=================   ПОЛУЧАЕМ СПИСОК СЕКЦИЙ ИЗ INI-ФАЙЛА   ====================}
{==============================================================================}
function TFMAIN.GetListTabIni(const PList: PStringList; const SPathIni: String): Boolean;
var FIni : TIniFile;
    S    : String;
begin
    {Инициализация}
    Result:=false;
    If PList=nil then Exit;
    PList^.Clear;
    If Not FileExists(SPathIni) then Exit;

    {Список таблиц}
    FIni:=TIniFile.Create(SPathIni);
    try     S := FIni.ReadString(INI_PARAM, INI_PARAM_TABLES, '');
            If Not SeparatMStr(S, PList, ', ') then Exit;
    finally FIni.Free;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{=================   ПОЛУЧАЕМ СПИСОК СЕКЦИЙ ИЗ DAT-ФАЙЛА   ====================}
{==============================================================================}
function TFMAIN.GetListTabDat(const PList: PStringList; const SPathDat: String): Boolean;
var FIni : TMemIniFile;
begin
    {Инициализация}
    Result:=false;
    If PList=nil then Exit;
    PList^.Clear;
    If Not FileExists(SPathDat) then Exit;

    {Список таблиц}
    FIni:=TMemIniFile.Create(SPathDat);
    try     FIni.ReadSections(PList^);
    finally FIni.Free;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{===================   ЧИТАЕМ КЛЮЧИ СЕКЦИИ ИЗ DAT-ФАЙЛА   =====================}
{==============================================================================}
function TFMAIN.GetTabKey(const PList: PStringList; const SPathDat, Section: String): Boolean;
begin
    {Инициализация}
    Result:=false;
    If PList=nil then Exit;
    PList^.Clear;
    If Not FileExists(SPathDat) then Exit;

    {Список таблиц}
    Result := (ReadSectionIni(PList, SPathDat, Section) >= 0);
    If Not Result then PList^.Clear;
end;


{==============================================================================}
{============================   ЧИТАЕМ СЕКЦИЮ   ===============================}
{==============================================================================}
function TFMAIN.GetTabVal(const SPath, STab: String; const PList: PStringList): Boolean;
var FDat: TMemIniFile;
begin
    {Инициализация}
    Result:=false;
    If PList=nil then Exit;

    {Читаем данные}
    PList^.Clear;
    FDat:=TMemIniFile.Create(SPath);
    try     FDat.ReadSectionValues(STab, PList^);
    finally FDat.Free;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{==================           ЗАПИСЫВАЕМ СЕКЦИЮ             ===================}
{==============================================================================}
{==================   IsFullReplace - полная замена секции  ===================}
{==============================================================================}
function TFMAIN.SetTabVal(const SPath, STab: String; const PList: PStringList; const IsFullReplace: Boolean): Boolean;
var FDat    : TMemIniFile;
    S, SVal : String;
    I       : Integer;
begin
    {Инициализация}
    Result:=false;

    {Создаем каталоги}
    S := ExtractFilePath(SPath);
    If Not DirectoryExists(S) then If Not ForceDirectories(S) then Exit;

    FDat:=TMemIniFile.Create(SPath);
    try
       {При полной замене удаляем cекцию полностью}
       If IsFullReplace then FDat.EraseSection(STab);

       {Записываем данные}
       For I:=0 to PList^.Count-1 do begin
          S    := PList^.Names[I];
          SVal := PList^.Values[S];
          If (SVal='') or (SVal='0') then FDat.DeleteKey(STab, S)
                                     else FDat.WriteString(STab, S, SVal);
       end;

       {Секция должна существовать}
       If Not FDat.SectionExists(STab) then begin
          FDat.WriteString(STab, '0', '0');
          FDat.DeleteKey(STab, '0');
       end;

       FDat.UpdateFile;
    finally
       FDat.Free;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{======================   СОЗДАЕМ СЕКЦИЮ DAT-ФАЙЛА   ==========================}
{==============================================================================}
function TFMAIN.CreateTab(const SPath, STab: String): Boolean;
var FDat : TMemIniFile;
begin
    {Инициализация}
    Result:=false;

    {Создаем пустую секцию}
    FDat:=TMemIniFile.Create(SPath);
    try     FDat.WriteString(STab, '0', '0');
            FDat.DeleteKey(STab, '0');
            FDat.UpdateFile;
    finally FDat.Free;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{======================   ОЧИЩАЕМ СЕКЦИЮ DAT-ФАЙЛА   ==========================}
{==============================================================================}
function TFMAIN.ClearTab(const SPath, STab: String): Boolean;
begin
    {Инициализация}
    Result:=false;

    {Удаляем секцию}
    DelTab(SPath, STab);

    {Создаем пустую секцию}
    CreateTab(SPath, STab);

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{======================   УДАЛЯЕМ СЕКЦИЮ DAT-ФАЙЛА   ==========================}
{==============================================================================}
function TFMAIN.DelTab(const SPath, STab: String): Boolean;
var FDat : TMemIniFile;
begin
    {Инициализация}
    Result:=false;
    If Not FileExists(SPath) then begin Result:=true; Exit; end;

    {Удаляем секцию}
    FDat:=TMemIniFile.Create(SPath);
    try     FDat.EraseSection(STab);
            FDat.UpdateFile;
    finally FDat.Free;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


