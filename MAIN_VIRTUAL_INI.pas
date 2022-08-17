{******************************************************************************}
{*****************       ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ            *******************}
{*****************  ПО ИНИЦИАЛИЗАЦИИ ВИРТУАЛЬНЫХ МАССИВОВ   *******************}
{******************************************************************************}

{==============================================================================}
{============    ИНИЦИАЛИЗАЦИЯ СПИСКА ИНДЕКСОВ БАНКОВ ДАННЫХ    ===============}
{==============================================================================}
function TFMAIN.IniListBank: Boolean;
var Sr       : TSearchRec;
    S, SPath : String;
    I        : Integer;
    B        : TBank;

    procedure AddLBank(const Ind: Integer);
    begin
        SetLength(LBank, Length(LBank)+1);
        With LBank[Length(LBank)-1] do begin
           ID      := Ind;
           If ID=0 then Pref:='' else Pref:=IntToStr(ID);
           Caption := '';
           BD      := TADOConnection.Create(Self);
        end;
    end;
begin
    {Инициализация}
    Result := false;
    ClearListBank;
    If PATH_BD='' then Exit;
    If PATH_BD[Length(PATH_BD)]='\' then SPath := Copy(PATH_BD, 1, Length(PATH_BD)-1) // Отрезаем '\'
                                    else SPath := PATH_BD;

    {Поиск банков данных}
    try
       If FindFirst(SPath+'*', faDirectory, Sr) = 0 then begin
          repeat
             S:=Sr.Name;
             If (S='.') or (S='..') then Continue;
             Delete(S, 1, Length(PATH_BANK0));
             If S = ''          then AddLBank(0);
             If IsIntegerStr(S) then AddLBank(StrToInt(S));
          until FindNext(Sr) <> 0;
       end;
    finally
       FindClose(Sr);
    end;

    {Сортировка индексов банков данных}
    I := Length(LBank) - 1;
    While I > 0 do begin
        If LBank[I-1].ID > LBank[I].ID then begin
           B          := LBank[I-1];
           LBank[I-1] := LBank[I];
           LBank[I]   := B;
           I          := Length(LBank)-1;
        end else begin
           I := I - 1;
        end;
    end;

    {Результат}
    IMainBank := -1;
    Result    := Length(LBank) > 0;
end;



{==============================================================================}
{===================    ОЧИСТКА СПИСКА БАНКОВ ДАННЫХ    =======================}
{==============================================================================}
procedure TFMAIN.ClearListBank;
var I: Integer;
begin
    try
       For I:=0 to Length(LBank)-1 do begin
          CloseBD(@LBank[I].BD, []);
          LBank[I].BD.Free;
       end;
    finally
       SetLength(LBank, 0);
    end;
end;


{==============================================================================}
{====================    ИНИЦИАЛИЗАЦИЯ ИНДЕКСНЫХ ФАЙЛОВ    ====================}
{==============================================================================}
function TFMAIN.IniIndexFile: Boolean;
var SSrc, SDect : String;
    I           : Integer;
begin
    {Инициализация}
    Result := false;
    try
       For I:=0 to Length(LBank)-1 do begin
          {SDect - имя индексного файла}
          SSrc := PATH_PROG+PATH_BANK0+LBank[I].Pref+'\'+FILE_LOCAL_INDEX;
          If Not FileExists(SSrc) then begin ErrMsg('Не найден индексный файл!'+CH_NEW+SSrc); Exit; end;
          If IS_NET then begin
             {Для сети копируем индексный файл на компьютер}
             SDect := PATH_WORK_TEMP+ExtractFileNameWithoutExt(FILE_LOCAL_INDEX)+LBank[I].Pref+'.mdb';
             If FileExists(SDect) then If Not DeleteFile(SDect)  then begin ErrMsg('Ошибка удаления старого индексного файла: '+CH_NEW+SDect); Exit; end;
             If Not CopyFileTo(SSrc, SDect)                      then begin ErrMsg('Ошибка копирования индексного файла: '     +CH_NEW+SDect); Exit; end;
             try FileSetAttr(SDect, 0); except end; {Убираем возможные атрибуты !!!!! ПРИ ЛЮБОМ ИЗМЕНЕНИИ ТРЕБУЕТСЯ ПОЛНАЯ ПЕРЕКОМПИЛЯЦИЯ С КОМПОНЕНТАМИ !!!!!}
          end else begin
             SDect := SSrc;
          end;

         {Открываем индексный файл}
         If Not OpenBD(@LBank[I].BD, SDect, '', [], []) then begin ErrMsg('Ошибка открытия индексного файла: '+CH_NEW+SDect); Exit; end;
       end;
       Result := true;
    finally
    end;
end;


{==============================================================================}
{=====================    ИНИЦИАЛИЗАЦИЯ ФАЙЛА ПОМОЩИ    =======================}
{==============================================================================}
function TFMAIN.IniHelpFile: Boolean;
var SSrc, SDect : String;
begin
    {Инициализация}
    Result:=false;
    try
       {SDect - имя файла помощи}
       If IS_ADMIN then SSrc := PATH_PROG + FILE_LOCAL_HELP_ADMIN
                   else SSrc := PATH_PROG + FILE_LOCAL_HELP_USER;
       If IS_NET then begin
          {Для сети копируем файл помощи на компьютер}
          If Not FileExists(SSrc) then Exit;
          SDect := PATH_WORK_TEMP+ExtractFileName(SSrc);
          If FileExists(SDect) then If Not DeleteFile(SDect) then begin ErrMsg('Ошибка удаления старого файла помощи: '+CH_NEW+SDect); Exit; end;
          If Not CopyFileTo(SSrc, SDect)                     then begin ErrMsg('Ошибка копирования файла помощи: '     +CH_NEW+SDect); Exit; end;
          try FileSetAttr(SDect, 0); except end; {Убираем возможные атрибуты !!!!! ПРИ ЛЮБОМ ИЗМЕНЕНИИ ТРЕБУЕТСЯ ПОЛНАЯ ПЕРЕКОМПИЛЯЦИЯ С КОМПОНЕНТАМИ !!!!!}
       end else begin
          SDect := SSrc;
       end;

       {Подключаем справку}
       HelpFile    := SDect;
       HelpType    := htKeyWord;
       HelpKeyWord := 'Set';

       Result := true;
    finally
    end;
end;


{==============================================================================}
{====================    ИНИЦИАЛИЗАЦИЯ СПИСКА РЕГИОНОВ    =====================}
{==============================================================================}
function TFMAIN.IniListRegions: Boolean;
var T           : TADOTable;
    IReg, IBank : Integer;

    function FindChain(const Ind: Integer): Boolean;
    var I: Integer;
    begin
        {Инициализация}
        Result := (Ind = IMAIN_REGION);
        If Ind <= IMAIN_REGION then Exit;
        I:=Ind;
        While (I > 0) and (Not Result) do begin
           I:=RegionsCounterToParent(I);
           If I=IMAIN_REGION then Result:=true;
        end;
    end;

begin
    {Инициализация}
    Result := false;
    IReg  := 0;
    SetLength(LRegion, 0);

    {*** Инициализация главного региона ***************************************}
    IMAIN_REGION := ReadGlobalInteger(INI_SET, INI_SET_IMAIN_REGION, 0);
    If IMAIN_REGION=0 then begin
       MessageDlg('Задайте главный регион!', mtWarning, [mbOk], 0);
       ASetExecute(nil);
       Exit;
    end;

    {*** Инициализация списка регионов ****************************************}
    T := TADOTable.Create(Self);
    try
       For IBank:=Low(LBank) to High(LBank) do begin
          If Not FindBDTables(@LBank[IBank].BD, [TABLE_REGIONS]) then Continue;
          If Not OpenBD(@LBank[IBank].BD, '', '', [@T], [TABLE_REGIONS]) then begin ErrMsg('Ошибка открытия таблицы: '+TABLE_REGIONS); Exit; end;

          SetLength(LRegion, T.RecordCount);
          T.Sort := '['+F_PARENT+'] ASC, ['+F_COUNTER+'] ASC';   // Если сортировать только по индексам, то при объединении регионов необходимо менять их индексы, что не правильно
          T.First;
          While Not T.Eof do begin
             {Если главный регион}
             If (IReg = 0) and (T.Fields[NREGIONS_COUNTER].AsInteger = IMAIN_REGION) then begin
                LRegion[0].FCounter := T.Fields[NREGIONS_COUNTER].AsInteger;
                LRegion[0].FCaption := T.Fields[NREGIONS_CAPTION].AsString;
                LRegion[0].FParent  := 0;
                LRegion[0].FGroup   := T.Fields[NREGIONS_GROUP].AsInteger;
                LRegion[0].FBegin   := T.Fields[NREGIONS_BEGIN].AsDateTime;
                LRegion[0].FEnd     := T.Fields[NREGIONS_END].AsDateTime;
                IReg := IReg + 1;
                T.Next;
                Continue;
             end;
             {Допустимость подчиненного региона}
             If FindChain(T.Fields[NREGIONS_PARENT].AsInteger) then begin
                LRegion[IReg].FCounter := T.Fields[NREGIONS_COUNTER].AsInteger;
                LRegion[IReg].FCaption := T.Fields[NREGIONS_CAPTION].AsString;
                LRegion[IReg].FParent  := T.Fields[NREGIONS_PARENT].AsInteger;
                LRegion[IReg].FGroup   := T.Fields[NREGIONS_GROUP].AsInteger;
                LRegion[IReg].FBegin   := T.Fields[NREGIONS_BEGIN].AsDateTime;
                LRegion[IReg].FEnd     := T.Fields[NREGIONS_END].AsDateTime;
                IReg := IReg + 1;
             end;
             T.Next;
          end;
          Break;
       end;
       If Length(LRegion) <> IReg then SetLength(LRegion, IReg);
       Result := Length(LRegion) > 0;
     finally
       CloseBD(nil, [@T]);
       T.Free;
    end;
    If Not Result then begin ErrMsg('Не подключена таблица: '+TABLE_REGIONS); Exit; end;

    {*** Имя главного региона *************************************************}
    SMAIN_REGION := RegionsCounterToCaption(IMAIN_REGION);
    If SMAIN_REGION='' then begin ErrMsg('Ошибка определения имени главного региона ['+IntToStr(IMAIN_REGION)+'] !'); Result:=false; end;
    StatusBar.Panels[STATUS_REGION].Text := ' [ '+SMAIN_REGION+' ] ';
end;


{==============================================================================}
{===========   ИНИЦИАЛИЗАЦИЯ СПИСКА ТАБЛИЦ, СТРОК И СТОЛБЦОВ   ================}
{==============================================================================}
{=====  ВАЖНО: В СПИСОК ТАБЛИЦ ПОМЕЩАЮТСЯ ВСЕ ТАБЛИЦЫ С УЧЕТОМ ПРИОРИТЕТА  ====}
{==============================================================================}
function TFMAIN.IniListTables(const IMin, IMax: Integer): Boolean;
var F_SPLASH            : TFSPLASH;
    TTabs, TRows, TCols : TADOTable;
    VTab                : TStringList;
    ITab, IRow, ICol    : Integer;
    IBank               : Integer;

    procedure SetProgress(const IVal: Integer);
    var DXBank, DXDop : Integer;
    begin
        DXBank := (IMax-IMin) Div Length(LBank);
        DXDop  := DXBank * IVal Div 100;
        DXBank := IMin + DXBank * (Length(LBank)-IBank-1) + DXDop;
        F_SPLASH.Msg(IntToStr(DXBank)+'%');
    end;

    procedure CopyRows(const IsAll: Boolean);
    var SInd : String;
    begin
        SetProgress(0);
        SetLength(LRow, Length(LRow) + TRows.RecordCount);
        TRows.First;
        While Not TRows.Eof do begin
           SInd := TRows.Fields[NROWS_TABLE].AsString;
           {Добавляем строку}
           If IsAll or (VTab.IndexOf(SInd)<0) then begin
              With LRow[IRow] do begin
                 FTable   := StrToInt(SInd);
                 FNumeric := TRows.Fields[NROWS_NUMERIC].AsString;
                 FCaption := TRows.Fields[NROWS_CAPTION].AsString;
              end;
              IRow := IRow + 1;
           end;
           {Следующая строка}
           TRows.Next;
        end;
        If Length(LRow) <> IRow then SetLength(LRow, IRow);
    end;

    procedure CopyCols(const IsAll: Boolean);
    var SInd : String;
    begin
        SetProgress(45);
        SetLength(LCol, Length(LCol) + TCols.RecordCount);
        TCols.First;
        While Not TCols.Eof do begin
           SInd := TCols.Fields[NCOLS_TABLE].AsString;
           {Добавляем столбец}
           If IsAll or (VTab.IndexOf(SInd)<0) then begin
              With LCol[ICol] do begin
                 FTable   := StrToInt(SInd);
                 FNumeric := TCols.Fields[NCOLS_NUMERIC].AsString;
                 FCaption := TCols.Fields[NCOLS_CAPTION].AsString;
              end;
              ICol := ICol + 1;
           end;
           {Следующий столбец}
           TCols.Next;
        end;
        If Length(LCol) <> ICol then SetLength(LCol, ICol);
    end;

    procedure CopyTables;
    var SInd : String;
    begin
        SetProgress(90);
        SetLength(LTable, Length(LTable) + TTabs.RecordCount);
        TTabs.First;
        While Not TTabs.Eof do begin
           SInd := TTabs.Fields[NTABLES_COUNTER].AsString;
           {Добавляем таблицу}
           //If IsAll or (VTab.IndexOf(SInd)<0) then begin
              With LTable[ITab] do begin
                 VTab.Add(SInd);
                 SBankPref := LBank[IBank].Pref;
                 FCounter  := SInd;
                 FCaption  := TTabs.Fields[NTABLES_CAPTION].AsString;
                 FNoStand  := TTabs.Fields[NTABLES_NOSTAND].AsBoolean;
                 FBegin    := TTabs.Fields[NTABLES_BEGIN].AsDateTime;
                 FEnd      := TTabs.Fields[NTABLES_END].AsDateTime;
                 If TTabs.Fields[NTABLES_CEP].AsBoolean then LCEP.Add(SInd);
              end;
              ITab := ITab + 1;
           //end;
           {Следующая таблица}
           TTabs.Next;
        end;
        If Length(LTable) <> ITab then SetLength(LTable, ITab);
    end;

begin
    {Инициализация}
    Result  := false;
    SetLength(LTable, 0); ITab := 0;
    SetLength(LRow,   0); IRow := 0;
    SetLength(LCol,   0); ICol := 0;
    F_SPLASH  := TFSPLASH(GlFindComponent('FSPLASH'));
    If F_SPLASH=nil then Exit;
    VTab  := TStringList.Create;
    TTabs := TADOTable.Create(Self);
    TRows := TADOTable.Create(Self);
    TCols := TADOTable.Create(Self);
    LCEP  := TStringList.Create;
    LCEP.Clear;
    try
       For IBank:=Low(LBank) to High(LBank) do begin
          If Not FindBDTables(@LBank[IBank].BD, [TABLE_TABLES, TABLE_ROWS, TABLE_COLS]) then Continue;
          If Not OpenBD(@LBank[IBank].BD, '', '', [@TTabs, @TRows, @TCols], [TABLE_TABLES, TABLE_ROWS, TABLE_COLS])
          then begin ErrMsg('Банк: '+IntToStr(IBank)+CH_NEW+'Ошибка открытия списков таблиц, строк, столбцов!'); Exit; end;

          {Создаем копии таблиц}
          CopyRows  (IBank = Low(LBank));
          CopyCols  (IBank = Low(LBank));
          CopyTables;                      // Таблицы копируются последними чтобы исключить столбцы и строки одной таблицы из разных банков

          CloseBD(nil, [@TRows, @TCols, @TTabs]);
       end;
       {Сортировка строк и столбцов}
       Result := (Length(LTable)>0) and (Length(LRow)>0) and (Length(LCol)>0);
    finally
       CloseBD(nil, [@TRows, @TCols, @TTabs]);
       TRows.Free; TCols.Free; TTabs.Free;
       VTab.Free;
    end;
    If Not Result then ErrMsg('Не подключены списки таблиц, строк, столбцов');
end;


(*  работает, но время выполнения выше на 0,5 ... 1,0 сек.  сортировка и т.д. сильно замедляет процесс
{==============================================================================}
{===========   ИНИЦИАЛИЗАЦИЯ СПИСКА ТАБЛИЦ, СТРОК И СТОЛБЦОВ   ================}
{==============================================================================}
function TFMAIN.IniListTables(const IMin, IMax: Integer): Boolean;
var F_SPLASH         : TFSPLASH;
    TTab, TRow, TCol : TADOTable;
    ITab, IRow, ICol : Integer;
    DSTab            : TDataSource;
    IBank            : Integer;
    STabInd          : String;

    procedure SetProgress;
    var DXBank, DXDop : Integer;
    begin
        DXBank := (IMax-IMin) Div Length(LBank);
        DXDop  := TTab.RecNo * DXBank Div TTab.RecordCount;
        DXBank := IMin + DXBank * (Length(LBank)-IBank-1) + DXDop;
        F_SPLASH.Msg(IntToStr(DXBank)+'%');
    end;

begin
    {Инициализация}
    Result  := false;
    SetLength(LTable, 0); ITab := 0;
    SetLength(LRow,   0); IRow := 0;
    SetLength(LCol,   0); ICol := 0;
    F_SPLASH  := TFSPLASH(GlFindComponent('FSPLASH'));
    If F_SPLASH=nil then Exit;
    TTab  := TADOTable.Create(Self);
    TRow  := TADOTable.Create(Self);
    TCol  := TADOTable.Create(Self);
    DSTab := TDataSource.Create(Self);
    DSTab.DataSet := TTab;
    LCEP  := TStringList.Create;
    LCEP.Clear;
    try
       For IBank:=Low(LBank) to High(LBank) do begin
          If Not FindBDTables(@LBank[IBank].BD, [TABLE_TABLES, TABLE_ROWS, TABLE_COLS]) then Continue;
          FreeTable(@TRow); FreeTable(@TCol);
          If Not OpenBD(@LBank[IBank].BD, '', '', [@TTab, @TRow, @TCol], [TABLE_TABLES, TABLE_ROWS, TABLE_COLS])
          then begin ErrMsg('Банк: '+IntToStr(IBank)+CH_NEW+'Ошибка открытия списков таблиц, строк, столбцов!'); Exit; end;
          SetLength(LTable, Length(LTable) + TTab.RecordCount);
          SetLength(LRow,   Length(LRow)   + TRow.RecordCount);
          SetLength(LCol,   Length(LCol)   + TCol.RecordCount);
          TTab.Sort := '['+F_COUNTER+'] ASC';   // Сортировка ОБЯЗАТЕЛЬНО до связи таблиц
          TRow.Sort := '['+F_NUMERIC+'] ASC';
          TCol.Sort := '['+F_NUMERIC+'] ASC';
          SetDBConnect(@TRow, @DSTab, F_COUNTER, TABLE_TABLES);
          SetDBConnect(@TCol, @DSTab, F_COUNTER, TABLE_TABLES);

          {Просматриваем таблицы}
          TTab.First;
          While Not TTab.Eof do begin
             STabInd := TTab.Fields[NTABLES_COUNTER].AsString;
             If Not FindTables(nil, true, STabInd, '', -1) then begin
                If TRow.RecordCount=0 then begin ErrMsg('Банк: '+IntToStr(IBank)+CH_NEW+'Нет строк для таблицы '   +STabInd+'!'); Exit; end;
                If TCol.RecordCount=0 then begin ErrMsg('Банк: '+IntToStr(IBank)+CH_NEW+'Нет столбцов для таблицы '+STabInd+'!'); Exit; end;
                SetProgress;

                {Добавляем таблицу}
                With LTable[ITab] do begin
                   SBankPref := LBank[IBank].Pref;
                   FCounter  := STabInd;
                   FCaption  := TTab.Fields[NTABLES_CAPTION].AsString;
                   FNoStand  := TTab.Fields[NTABLES_NOSTAND].AsBoolean;
                   If TTab.Fields[NTABLES_CEP].AsBoolean then LCEP.Add(FCounter);
                   FBegin    := TTab.Fields[NTABLES_BEGIN].AsDateTime;
                   FEnd      := TTab.Fields[NTABLES_END].AsDateTime;
                end;
                ITab := ITab+1;

                {Добавляем связанные строки}
                TRow.First;
                While Not TRow.Eof do begin
                    With LRow[IRow] do begin
                       FNumeric := TRow.Fields[NROWS_NUMERIC].AsString;
                       FCaption := TRow.Fields[NROWS_CAPTION].AsString;
                       FTable   := TRow.Fields[NROWS_TABLE].AsInteger;
                    end;
                    IRow := IRow+1;
                    TRow.Next;
                end;

                {Добавляем связанные столбцы}
                TCol.First;
                While Not TCol.Eof do begin
                    With LCol[ICol] do begin
                       FNumeric := TCol.Fields[NCOLS_NUMERIC].AsString;
                       FCaption := TCol.Fields[NCOLS_CAPTION].AsString;
                       FTable   := TCol.Fields[NCOLS_TABLE].AsInteger;
                    end;
                    ICol := ICol+1;
                    TCol.Next;
                end;
             end;
             TTab.Next;
          end;
          CloseBD(nil, [@TRow, @TCol, @TTab]);
       end;
       If Length(LTable) <> ITab then SetLength(LTable, ITab);
       If Length(LRow)   <> IRow then SetLength(LRow,   IRow);
       If Length(LCol)   <> ICol then SetLength(LCol,   ICol);
       Result := (Length(LTable)>0) and (Length(LRow)>0) and (Length(LCol)>0);
    finally
       CloseBD(nil, [@TRow, @TCol, @TTab]);
       TRow.Free; TCol.Free; TTab.Free;
       DSTab.Free;
    end;
    If Not Result then ErrMsg('Не подключены списки таблиц, строк, столбцов');
end;
*)

{==============================================================================}
{=====================   ИНИЦИАЛИЗАЦИЯ СПИСКА ФОРМ   ==========================}
{==============================================================================}
function TFMAIN.IniListForms: Boolean;
var T                 : TADOTable;
    Ind, IBank, IForm : Integer;
begin
    {Инициализация}
    Result := false;
    IForm  := 0;
    SetLength(LForma, 0);

    T := TADOTable.Create(Self);
    try
       For IBank:=Low(LBank) to High(LBank) do begin
          If Not FindBDTables(@LBank[IBank].BD, [TABLE_FORMS]) then Continue;
          If Not OpenBD(@LBank[IBank].BD, '', '', [@T], [TABLE_FORMS]) then begin ErrMsg('Ошибка открытия таблицы: '+TABLE_FORMS); Exit; end;
          SetLength(LForma, Length(LForma) + T.RecordCount);

          T.Sort := '['+F_COUNTER+'] ASC';
          T.First;
          While Not T.Eof do begin
             Ind := T.Fields[NFORMS_COUNTER].AsInteger;
             If Not FindForms(nil, true, false, Ind, -1, '', '', '', '', false, -1) then begin
                {Добавляем форму}
                With LForma[IForm] do begin
                   SBankPref   := LBank[IBank].Pref;
                   FCounter    := Ind;
                   FParent     := T.Fields[NFORMS_PARENT].AsInteger;
                   FIcon       := T.Fields[NFORMS_ICON].AsInteger;
                   FExport     := T.Fields[NFORMS_EXPORT].AsBoolean;
                   FRegion     := T.Fields[NFORMS_REGION].AsString;
                   FCaption    := Trim(T.Fields[NFORMS_CAPTION].AsString);
                   FFile       := Trim(T.Fields[NFORMS_FILE].AsString);
                   FMonth1     := T.Fields[NFORMS_MONTH1].AsBoolean;
                   FMonth3     := T.Fields[NFORMS_MONTH3].AsBoolean;
                   FBegin      := T.Fields[NFORMS_BEGIN].AsDateTime;
                   FEnd        := T.Fields[NFORMS_END].AsDateTime;
                   FCellID     := Trim(T.Fields[NFORMS_CELL_ID].AsString);
                   FStrID      := Trim(T.Fields[NFORMS_STR_ID].AsString);
                   FCellPeriod := Trim(T.Fields[NFORMS_CELL_PERIOD].AsString);
                   FCellRegion := Trim(T.Fields[NFORMS_CELL_REGION].AsString);
                end;
                IForm := IForm + 1;
             end;
             T.Next;
          end;
          CloseBD(nil, [@T]);
       end;
       If Length(LForma) <> IForm then SetLength(LForma, IForm);
       Result := Length(LForma)>0;
    finally
       CloseBD(nil, [@T]);
       T.Free;
    end;
    If Not Result then ErrMsg('Не подключена таблица: '+TABLE_FORMS); 
end;


{==============================================================================}
{===============   ИНИЦИАЛИЗАЦИЯ СПИСКА АЛЬТЕРНАТИВНЫХ КОДОВ   ================}
{==============================================================================}
function TFMAIN.IniORKey: Boolean;
var T           : TADOTable;
    IBank, IKey : Integer;
    S1, S2 : String;
begin
    {Инициализация}
    Result := false;
    IKey   := 0;
    SetLength(LORKey, 0);

    T := TADOTable.Create(Self);
    try
       For IBank:=Low(LBank) to High(LBank) do begin
          If Not FindBDTables(@LBank[IBank].BD, [TABLE_ORKEY]) then Continue;
          If Not OpenBD(@LBank[IBank].BD, '', '', [@T], [TABLE_ORKEY]) then begin ErrMsg('Ошибка открытия таблицы: '+TABLE_ORKEY); Exit; end;
          SetLength(LORKey, Length(LORKey) + T.RecordCount);

          T.First;
          While Not T.Eof do begin
             S1 := T.Fields[NORKEY_1].AsString;
             S2 := T.Fields[NORKEY_2].AsString;
             If (Not FindORKeys(nil, true, S1, S2)) and (Not FindORKeys(nil, true, S2, S1)) then begin
                With LORKey[IKey] do begin
                   FKey1 := S1;
                   FKey2 := S2;
                end;
                IKey := IKey + 1;
             end;
             T.Next;
          end;
          CloseBD(nil, [@T]);
       end;
       If Length(LORKey) <> IKey then SetLength(LORKey, IKey);
       Result := true;
    finally
       CloseBD(nil, [@T]);
       T.Free;
    end;
    //If Not Result then ErrMsg('Не подключена таблица: '+TABLE_ORKEY);
end;

