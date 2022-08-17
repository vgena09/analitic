{******************************************************************************}
{************************  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ  ***************************}
{************************   ПО РАБОТЕ С РЕГИОНАМИ   ***************************}
{******************************************************************************}


{==============================================================================}
{=====================   ДОПУСТИМОСТЬ ИЗМЕНЕНИЙ РЕГИОНА   =====================}
{==============================================================================}
{===========             ПЕРИОД РЕГИОНА НЕ ПРОВЕРЯЕТСЯ                  =======}
{==============================================================================}
{===========  Если регион не имеет подчиненных субрегионов:             =======}
{===========  LTab - список предлагаемых к созданию/модификации таблиц  =======}
{===========  LTab = nil - все имеющиеся таблицы региона                =======}
{==============================================================================}
function TFMAIN.IsModifyRegion(const SYear, SMonth, SRegion: String; const PLTab: PStringList): Boolean;
var LRegion_           : TLRegion;
    SRegion1, SRegion2 : String;
    FDate              : TDate;
    I, Ind             : Integer;

    LTabSub            : TStringList;
    LFind              : array of String;
begin
    {Инициализация}
    Result   := false;
    If (SYear='') or (SMonth='') or (PLTab=nil) then Exit;
    SRegion1 := ReplModulStr(SRegion, '', SUB_RGN1, SUB_RGN2);
    SRegion2 := CutModulStr (SRegion,     SUB_RGN1, SUB_RGN2);
    FDate    := SDatePeriod(SYear, SMonth);

    try
       {Региона в списке нет -  модификация запрещена}
       If Not FindRegions(@LRegion_, true, -1, SRegion1, -1, -1, FDate) then Exit;
       Ind := LRegion_[0].FCounter;

       {Есть хоть одна стандартная таблица - проверка продолжается}
       Result := (PLTab^.Count > 0);
       For I:=0 to PLTab^.Count-1 do begin
          If Not IsTableNoStandart(PLTab^[I]) then begin
             Result:=false;
             Break;
          end;
       end;
       If Result then Exit;

       {Регион имеет подчиненные субрегионы - модификация запрещена}
       If FindRegions(@LRegion_, true, -1, '', Ind, -1, FDate) then Exit;
    finally
       SetLength(LRegion_, 0);
    end;

    {Свободный регион - модификация разрешена}
    If SRegion2<>'' then begin
       Result:=true;
       Exit;

    {Регион не имеет подчиненных субрегионов}
    end else begin
       try
          {Регион имеет свободные субрегионы - просматриваем банки данных}
          If FindDataYMR(@LFind, SYear+'\'+SMonth+'\'+SRegion1+SUB_RGN1+'*', true) > 0 then begin
             {Инициализация}
             If PLTab^.Count=0 then Exit;
             LTabSub := TStringList.Create;
             try
                {Просматриваем все свободные субрегионы}
                For I:=Low(LFind) to High(LFind) do begin
                   {Получаем список таблиц свободного субрегиона}
                   If Not GetListTabDat(@LTabSub, LFind[I]) then Exit;
                   {Свободный субрегион не должен содержать ни одной из заданных таблиц}
                   For Ind:=0 to PLTab^.Count-1 do begin
                      If LTabSub.IndexOf(PLTab^[Ind])>=0 then Exit;
                   end;
                end;
                Result:=true;
             finally
                LTabSub.Free;
             end;

          {Регион не имеет свободных субрегионов - модификация разрешена}
          end else begin
             Result:=true;
          end;
       finally
          SetLength(LFind, 0);
       end;
    end;
end;


{==============================================================================}
{======================  СПИСОК РЕГИОНОВ ДЛЯ АНАЛИТИКИ  =======================}
{==============================================================================}
function TFMAIN.ListRegionsAnalitic(const PList: PStringList; const SBaseRegion: String;
                                    const IsApparat, IsMainRegion: Boolean;
                                    const SYear, SMonth: String): Boolean;
var LRegion_ : TLRegion;
    FDate    : TDate;
    I, Ind   : Integer;
begin
    {Инициализация}
    Result := false;
    If PList=nil then Exit;
    FDate  := SDatePeriod(SYear, SMonth);
    PList^.Clear;
    try
       {Индекс главного региона}
       If Not FindRegions(@LRegion_, true, -1, SBaseRegion, -1, -1, FDate) then begin ErrMsg('Ошибка определения индекса региона ['+SBaseRegion+'] !'); Exit; end;
       Ind := LRegion_[0].FCounter;

       {Список зависимых регионов}
       If FindRegions(@LRegion_, false, -1, '', Ind, -1, FDate) then begin
          For I:=1 to Length(LRegion_)-1 do PList^.Add(LRegion_[I].FCaption);
       end;

       {Добавляем аппарат}
       If IsApparat and (Length(LRegion_)>0) then PList^.Add(LRegion_[0].FCaption);

       {Добавляем главный регион}
       If IsMainRegion then PList^.Add(SBaseRegion);

       {Возвращаемый результат}
       Result:=true;

    finally
       SetLength(LRegion_, 0);
    end;
end;


{==============================================================================}
{===========  ИЕРАРХИЧЕКИЙ ВОСХОДЯЩИЙ СПИСОК ЗАВИСИМЫХ РЕГИОНОВ   =============}
{==============================================================================}
{===========    Браславский [...], Браславский, Витебский, РБ     =============}
{==============================================================================}
{===========         Если SYEAR, SMONTH = '' - все регионы        =============}
{==============================================================================}
function TFMAIN.ListRegionsRangeUp(const PList: PStringList; const SBaseRegion: String;
                                   const SYear, SMonth: String): Boolean;
var LRegion_ : TLRegion;
    FDate    : TDate;
    S1, S2   : String;
    IParent  : Integer;
begin
    {Инициализация}
    Result := false;
    If PList=nil then Exit;
    PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);
    S2    := CutModulStr(SBaseRegion, SUB_RGN1, SUB_RGN2);
    S1    := ReplModulStr(SBaseRegion, '', SUB_RGN1, SUB_RGN2);
    If S1='' then Exit;

    try
       {Если региона в списке нет}
       If Not FindRegions(@LRegion_, true, -1, S1, -1, -1, FDate) then Exit;
       IParent := LRegion_[0].FParent;

       {Cубрегион}
       If S2<>'' then PList^.Add(SBaseRegion);

       {Базовый регион}
       PList^.Add(S1);

       {Родительские регионы}
       While IParent>0 do begin
          If Not FindRegions(@LRegion_, true, IParent, '', -1, -1, FDate) then Exit;
          PList^.Add(LRegion_[0].FCaption);
          IParent := LRegion_[0].FParent;
       end;
    finally
       SetLength(LRegion_, 0);
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{===========   ИЕРАРХИЧЕКИЙ НИСХОДЯЩИЙ СПИСОК ЗАВИСИМЫХ РЕГИОНОВ   ============}
{==============================================================================}
{===========         Витебский, Браславский, Полоцкий, ...         ============}
{==============================================================================}
{===========         Если SYEAR, SMONTH = '' - все регионы         ============}
{==============================================================================}
function TFMAIN.ListRegionsRangeDown(const PList: PStringList; const SBaseRegion: String;
                                     const SYear, SMonth: String): Boolean;
var LRegion_ : TLRegion;
    FDate    : TDate;
    S1, S2   : String;

    procedure AddReg(const IParent: Integer);
    var LRegion_ : TLRegion;
        I        : Integer;
    begin
        try
           {Получаем список зависимых регионов}
           If Not FindRegions(@LRegion_, false, -1, '', IParent, -1, FDate) then Exit;
           {Просматриваем список}
           For I:=0 to Length(LRegion_)-1 do begin
              PList^.Add(LRegion_[I].FCaption);
              AddReg(LRegion_[I].FCounter);
           end;
        finally
           SetLength(LRegion_, 0);
        end;
    end;

begin
    {Инициализация}
    Result := false;
    If PList=nil then Exit;
    PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);
    S2    := CutModulStr(SBaseRegion,      SUB_RGN1, SUB_RGN2);
    S1    := ReplModulStr(SBaseRegion, '', SUB_RGN1, SUB_RGN2);
    If (S1='') or (S2<>'') then Exit;

    try
       {Если региона в списке нет}
       If Not FindRegions(@LRegion_, true, -1, S1, -1, -1, FDate) then Exit;

       {Рекурсия начиная с главного региона}
       PList^.Add(LRegion_[0].FCaption);
       AddReg(LRegion_[0].FCounter);
    finally
       SetLength(LRegion_, 0);
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{======================  СПИСОК РЕГИОНОВ ОДНОГО УРОВНЯ  =======================}
{==============================================================================}
function TFMAIN.ListRegionsLevel(const PList: PStringList; const SBaseRegion: String;
                                 const SYear, SMonth: String): Boolean;
var LRegion_  : TLRegion;
    FDate     : TDate;
    I, ILevel : Integer;
begin
    {Инициализация}
    Result := false;
    If PList=nil then Exit;                PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);   If FDate=0 then Exit;
    If IsSubRegion(SBaseRegion) then Exit;

    try
       {Уровень главного региона}
       If Not FindRegions(@LRegion_, true, -1, SBaseRegion, -1, -1, FDate) then Exit;
       ILevel := LRegion_[0].FGroup;

       {Выбираем регионы с одним уровнем}
       If Not FindRegions(@LRegion_, false, -1, '', -1, ILevel, FDate) then Exit;

       {Формируем список регионов}
       For I:=Low(LRegion_) to High(LRegion_) do PList^.Add(LRegion_[I].FCaption);
    finally
       SetLength(LRegion_, 0);
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{==========================   КОПИЯ СПИСКА РЕГИОНОВ   =========================}
{==============================================================================}
function TFMAIN.CopyListRegions(const PList: PStringList; const SYear, SMonth: String): Boolean;
var FDate : TDate;
    I     : Integer;
begin
    {Инициализация}
    Result := false;
    If PList=nil then Exit;
    PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);
    If FDate=0 then Exit;

    {Заполняем список}
    For I:=Low(LRegion) to High(LRegion) do begin
       If LRegion[I].FBegin > 0 then If LRegion[I].FBegin > FDate then Continue;
       If LRegion[I].FEnd   > 0 then If LRegion[I].FEnd   < FDate then Continue;
       PList^.Add(LRegion[I].FCaption);
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{=========================    СОЗДАЕТ DAT-ФАЙЛЫ    ============================}
{=========================    ПУТЬ ДЛЯ TREEPATH    ============================}
{==============================================================================}
function TFMAIN.CreateRegions(const SYear, SMonth, SRegion: String): String;
var SList   : TStringlist;
    SPath   : String;
    SFile   : String;
    Result0 : String;
    I       : Integer;
begin
    {Инициализация}
    Result  := '';
    Result0 := '';
    SPath   := PATH_BD_DATA+SYear+'\'+SMonth+'\';
    SList   := TStringList.Create;
    try
       {Получаем иерархический восходящий список регионов}
       If Not ListRegionsRangeUp(@SList, SRegion, SYear, SMonth) then Exit;
       If SList.Count=0 then Exit;

       {Просматриваем регионы}
       For I:=0 to SList.Count-1 do begin
          Result0 := SList[I]+'\'+Result0;
          SFile   := SPath+SList[I]+'.dat';
          If Not CreateEmptyFile(SFile) then begin ErrMsg('Ошибка создания файла ['+SFile+']!'); Exit; end;
       end;

    finally
       SList.Free;
    end;

    {Возвращаемый результат}
    Result:=SYear+'\'+SMonth+'\'+Result0;
end;


{==============================================================================}
{========   ПРОВЕРЯЕТ ДОПУСТИМОСТЬ НАЗВАНИЯ РЕГИОНА ИЛИ СУБРЕГИОНА    =========}
{==============================================================================}
function TFMAIN.VerifyRegionName(const SRegionName: String): Boolean;
var S : String;
    I : Integer;
begin
    {Инициализация}
    Result := false;
    {Выделяем название региона}
    S := ReplModulStr(SRegionName, '', SUB_RGN1, SUB_RGN2);

    {Проверяем название региона}
    For I:=Low(LRegion) to High(LRegion) do begin
       If LRegion[I].FCaption = S then begin
          Result:=true;
          Break;
       end;
    end;
end;


{==============================================================================}
{================   ЯВЛЯЕТСЯ ЛИ НАЗВАНИЕ РЕГИОНА СУБРЕГИОНОМ    ===============}
{==============================================================================}
function TFMAIN.IsSubRegion(const SRegion: String): Boolean;
begin
    Result := (Pos(SUB_RGN1, SRegion) > 0);
end;


(*
{==============================================================================}
{===========================   УРОВЕНЬ РЕГИОНА   ==============================}
{==============================================================================}
function TFMAIN.RegionLevel(const SRegion: String; const FDate: TDate): Integer;
var LRegion_ : TLRegion;
    IParent  : Integer;
    IResult  : Integer;
begin
    {Инициализация}
    Result  := -1;
    IResult := -1;
    If FDate=0 then Exit;
    If IsSubRegion(SRegion) then Exit;
    SetLength(LRegion_, 0);
    try
       {Индекс родителя главного региона}
       If Not FindRegions(@LRegion_, true, -1, SRegion, -1, -1, FDate) then begin ErrMsg('Ошибка определения индекса родительского региона ['+SRegion+'] !'); Exit; end;
       IParent  := LRegion_[0].FParent;
       IResult  := 0;
       {Просматриваем цепочку}
       While IParent > 0 do begin
          If Not FindRegions(@LRegion_, true, IParent, '', -1, -1, FDate) then begin ErrMsg('Ошибка определения индекса родительского региона ['+SRegion+'] !'); Exit; end;
          IParent  := LRegion_[0].FParent;
          IResult  := IResult + 1;
       end;
       Result := IResult;
    finally
       SetLength(LRegion_, 0);
    end;
end;
*)

