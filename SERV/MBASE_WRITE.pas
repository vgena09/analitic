{==============================================================================}
{========================   ЗАПИСЫВАЕМ ЗНАЧЕНИЕ   =============================}
{==============================================================================}
function TBASE.SetValYMR(const YMR, STab, SCol, SRow, SVal: String): Boolean;
var SPath : String;
begin
    SPath  := FFMAIN.TabToDatYMR(STab, YMR);
    Result := SetVal(SPath, STab, SCol, SRow, SVal);
end;


{==============================================================================}
{========================   ЗАПИСЫВАЕМ ЗНАЧЕНИЕ   =============================}
{==============================================================================}
function TBASE.SetVal(const SPath, STab, SCol, SRow, SVal: String): Boolean;
var SVal0 : String;
    PFile : PMemIniFile;
begin
    {Инициализация}
    Result := false;
    SVal0  := SVal;

    {Определяем файл в памяти}
    PFile := FindMemFile(SPath);
    If PFile=nil then begin
       PFile := AddMemFile(SPath);
       If PFile=nil then begin ErrMsg('Ошибка подключения к dat-файлу:'+CH_NEW+SPath); Exit; end;
    end;

    try
       {Если дробь, то округляем до целого}
       If (Pos(',', SVal0)>0) or (Pos('.', SVal0)>0) then begin
          SVal0 := ReplStr(SVal0, '.', ',');
          try    SVal0 := IntToStr(Round(StrToFloat(SVal0)));
          except SVal0 := '';
          end;
       end;

       {Записываем данные}
       If (SVal0='') or (SVal0='0') or (Not IsIntegerStr(SVal0)) then PFile^.DeleteKey  (STab, SCol+' '+SRow)
                                                                 else PFile^.WriteString(STab, SCol+' '+SRow, SVal0);

       {Секция должна существовать}
       If Not PFile^.SectionExists(STab) then begin
          PFile^.WriteString(STab, '0', '0');
          PFile^.DeleteKey  (STab, '0');
       end;

       {Возвращаемый результат}
       Result:=true;
    finally
    end;
end;


{==============================================================================}
{===================       ПОИСК DAT-ФАЙЛА В КЭШЕ       =======================}
{==============================================================================}
{===================   SPath  - путь к искомому файлу   =======================}
{===================   RESULT - индекс кэша             =======================}
{==============================================================================}
function TBASE.FindMemFile(const SPath: String): PMemIniFile;
var I      : Integer;
    SPath0 : String;
begin
    Result := nil;
    SPath0 := AnsiUpperCase(SPath);
    For I:=Low(MemFile) to High(MemFile) do begin
       If AnsiUpperCase(MemFile[I].FileName) = SPath0 then begin
          Result := @MemFile[I];
          Break;
       end;
    end;
end;


{==============================================================================}
{========================  ОТКРЫТИЕ MEM-ФАЙЛА  ================================}
{==============================================================================}
function TBASE.AddMemFile(const SPath: String): PMemIniFile;
var I : Integer;
begin
    Result := nil;
    try
       I := Length(MemFile);
       SetLength(MemFile, I + 1);
       MemFile[I] := TMemIniFile.Create(SPath);
       Result := @MemFile[I];
    finally
    end;
end;


{==============================================================================}
{========  СОХРАНЕНИЕ ВСЕГО КЭША В DAT-ФАЙЛЫ С ПОЛНЫМ ИХ ЗАМЕЩЕНИЕМ  ==========}
{==============================================================================}
procedure TBASE.SaveMemFile;
var I : Integer;
begin
    For I:=Low(MemFile) to High(MemFile) do MemFile[I].UpdateFile;
end;

