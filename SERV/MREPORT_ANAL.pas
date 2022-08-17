{==============================================================================}
{====================     Формирует область ввода    ==========================}
{==============================================================================}
{====================   SPage  = 1, 5                ==========================}
{====================   SBlock = 'A5:BH125'          ==========================}
{====================   SBlock = S1+':'+S2           ==========================}
{====================   SBlock = Row1,Col1,Row2,Col2 ==========================}
{==============================================================================}
function TREPORT.CreateBlockAnal: Boolean;
label Dl;
var Row1, Col1    : Integer;
    Row2, Col2    : Integer;
    Ind,  ICol    : Integer;
    S, S1, S2, S3 : String;
    AData, Range  : Variant;
    IsKod         : Boolean;    // Признак обработки кода
    LParamNew     : TListParam; // для подмены (отчет: просмотр): список параметров для подмены (отчет: просмотр)
    PParamOld     : PListParam; // для подмены (отчет: просмотр)

    {Номер первого столбца данных}
    function FirstColDat: Integer;
    var ICol: Integer;
    begin
        Result := Col1;
        For ICol := Col1+1 to Col1 + 10 do begin
           If CmpStr(FIni.ReadString(INI_PAGE+SPage, ExcelColumnNumToChar(ICol)+IntToStr(Row1), ''), INI_PAGE_COL_KOD) then begin
              Result:=ICol;
              Break;
           end;
        end;
    end;

    {$INCLUDE MREPORT_ANAL_COL}

begin
    {Инициализация}
    Result := false;
    IsKod  := false;
    If (CmpStrList(SType, [INI_PAGE_TYPE_STANDART]) < 0) and (Length(Analiz.PLParam^) = 0) then Exit;
    S1 := AnsiUpperCase(CutSlovo(SBlock, 1, ':'));
    S2 := AnsiUpperCase(CutSlovo(SBlock, 2, ':'));
    If S1='' then Exit;
    If S2='' then S2 := S1;
    If Not ExcelSeparatAddress(S1, Col1, Row1) then Exit;
    If Not ExcelSeparatAddress(S2, Col2, Row2) then Exit;

    SetInfo('настройка таблиц');
    Analiz.IStep          := FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_STEP, 12);
    Analiz.ILoop          := FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_LOOP,  2);

    {При необходимости делаем подмену параметров}
    SetLength(LParamNew, 0);
    PParamOld := Analiz.PLParam;
    Ind       := 0;
    S         := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_STAT_PARAM+IntToStr(Ind+1), '');
    While S <> '' do begin
       Inc(Ind); SetLength(LParamNew, Ind);
       LParamNew[Ind-1].SCaption := Trim(TokChar(S, '|'));
       LParamNew[Ind-1].SFormula := Trim(TokChar(S, '|'));
       LParamNew[Ind-1].SColor   := Trim(TokChar(S, '|'));
       If LParamNew[Ind-1].SCaption = '' then LParamNew[Ind-1].SCaption := FFMAIN.FormulaToCaption(LParamNew[Ind-1].SFormula, true);
       S := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_STAT_PARAM+IntToStr(Ind+1), '');
    end;
    If Length(LParamNew) > 0 then Analiz.PLParam := @LParamNew;

    {===  Отчет: стандартный  =================================================}
    If CmpStrList(SType, [INI_PAGE_TYPE_STANDART])>=0 then begin
       Analiz.IRowBefore  := 0;
       Analiz.IRowAfter   := 2;
       Analiz.FirstColDat := Col1;
       {Параметры аналитики}
       FFMAIN.ListRegionsAnalitic(@Analiz.LRegions, SRegion, true, true, SYear, SMonth);
       Row2 := Row1-1+(Analiz.LRegions.Count*Analiz.ILoop)+Analiz.IRowAfter;
       {Размеры таблицы: по вертикали}
       If Not FIni.ReadBool(INI_PAGE+SPage, INI_PAGE_FIXSIZETABLE, false) then begin
          If Analiz.LRegions.Count > 1 then begin
             WB.WorkSheets[IPage].Rows[IntToStr(Row1)+':'+IntToStr(Row1-1+Analiz.ILoop)].Copy;
             WB.WorkSheets[IPage].Rows[IntToStr(Row1+Analiz.ILoop)+':'+IntToStr(Row2-Analiz.IRowAfter-Analiz.ILoop)].Insert(xlDown, true);
          end else begin
             WB.WorkSheets[IPage].Rows[IntToStr(Row1)+':'+IntToStr(Row1+1)].Delete;
          end;
       end;
       Goto Dl;

    {===  Отчет: нестандартный  ===============================================}
    end else begin
       Analiz.IRowBefore  := FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_BEFORE, 0);
       Analiz.IRowAfter   := FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_AFTER,  0);
       Analiz.FirstColDat := FirstColDat;
    end;

    {===  Отчет: сравнение  ===================================================}
    If CmpStrList(SType, [INI_PAGE_TYPE_COMPARE])>=0 then begin
       {Параметры аналитики}
       FFMAIN.ListRegionsAnalitic(@Analiz.LRegions, SRegion, true, true, SYear, SMonth);
       Col2 := Analiz.FirstColDat-1+Length(Analiz.PLParam^);
       Row2 := Row1-1+(Analiz.LRegions.Count*Analiz.ILoop)+Analiz.IRowAfter;
       {Размеры таблицы: по горизонтали}
       If Length(Analiz.PLParam^) > 1 then begin
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat)].Copy;
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat+1)+':'+ExcelColumnNumToChar(Col2)].Insert(xlDown, true);
          WB.WorkSheets[IPage].Range[ExcelCombineAddress(Analiz.FirstColDat, Row1-Analiz.IRowBefore), ExcelCombineAddress(Col2, Row1-1+Analiz.ILoop+Analiz.ILoop+Analiz.IRowAfter)].Borders[xlInsideVertical].Weight := xlThin;
       end;
       {Размеры таблицы: по вертикали}
       If Analiz.LRegions.Count > 1 then begin
          WB.WorkSheets[IPage].Rows[IntToStr(Row1)+':'+IntToStr(Row1-1+Analiz.ILoop)].Copy;
          WB.WorkSheets[IPage].Rows[IntToStr(Row1+Analiz.ILoop)+':'+IntToStr(Row2-Analiz.IRowAfter-Analiz.ILoop)].Insert(xlDown, true);
       end else begin
          WB.WorkSheets[IPage].Rows[IntToStr(Row1)+':'+IntToStr(Row1+1)].Delete;
       end;
       Goto Dl;
    end;

    {===  Отчет: рейтинг  =====================================================}
    If CmpStrList(SType, [INI_PAGE_TYPE_RATING])>=0 then begin
       {Параметры аналитики}
       FFMAIN.ListRegionsRangeDown(@Analiz.LRegions, SRegion, SYear, SMonth);
       Col2 := Analiz.FirstColDat-1+Length(Analiz.PLParam^);
       Row2 := Row1-1+(Analiz.LRegions.Count*Analiz.ILoop)+Analiz.IRowAfter;
       {Размеры таблицы: по горизонтали}
       If Length(Analiz.PLParam^) > 1 then begin
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat)].Copy;
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat+1)+':'+ExcelColumnNumToChar(Col2)].Insert(xlDown, true);
          WB.WorkSheets[IPage].Range[ExcelCombineAddress(Analiz.FirstColDat, Row1-Analiz.IRowBefore), ExcelCombineAddress(Col2, Row1-1+Analiz.ILoop+Analiz.IRowAfter)].Borders[xlInsideVertical].Weight := xlThin;
       end;
       {Размеры таблицы: по вертикали}
       If Analiz.LRegions.Count > 1 then begin
          WB.WorkSheets[IPage].Rows[IntToStr(Row1)+':'+IntToStr(Row1-1+Analiz.ILoop)].Copy;
          WB.WorkSheets[IPage].Rows[IntToStr(Row1+Analiz.ILoop)+':'+IntToStr(Row2-Analiz.IRowAfter)].Insert(xlDown, true);
       end;
       Goto Dl;
    end;

    {===  Отчет: регион  ======================================================}
    If CmpStrList(SType, [INI_PAGE_TYPE_REGION])>=0 then begin
       {Параметры аналитики}
       FFMAIN.ListRegionsRangeUp(@Analiz.LRegions, SRegion, SYear, SMonth);
       Col2 := Analiz.FirstColDat-1+Analiz.LRegions.Count;
       Row2 := Row1-1+(Length(Analiz.PLParam^)*Analiz.ILoop);
       {Размеры таблицы: по горизонтали}
       If Analiz.LRegions.Count > 1 then begin
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat)].Copy;
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat+1)+':'+ExcelColumnNumToChar(Col2)].Insert(xlDown, true);
          WB.WorkSheets[IPage].Range[ExcelCombineAddress(Analiz.FirstColDat, Row1-Analiz.IRowBefore), ExcelCombineAddress(Col2, Row1-1+Analiz.ILoop+Analiz.IRowAfter)].Borders[xlInsideVertical].Weight := xlThin;
       end;
       {Размеры таблицы: по вертикали}
       If Length(Analiz.PLParam^) > 1 then begin
          WB.WorkSheets[IPage].Rows[IntToStr(Row1)+':'+IntToStr(Row1-1+Analiz.ILoop)].Copy;
          WB.WorkSheets[IPage].Rows[IntToStr(Row1+Analiz.ILoop)+':'+IntToStr(Row2-Analiz.IRowAfter)].Insert(xlDown, true);
       end;
       Goto Dl;
    end;

    {===  Отчет: динамика  ====================================================}
    If CmpStrList(SType, [INI_PAGE_TYPE_DYNAMIC])>=0 then begin
       {Параметры аналитики}
       Analiz.LRegions.Text := SRegion;
       Col2 := Analiz.FirstColDat-1+Length(Analiz.PLParam^);
       Row2 := Row1-1+Analiz.ILoop+Analiz.IRowAfter;
       {Размеры таблицы: по горизонтали}
       If Length(Analiz.PLParam^) > 1 then begin
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat)].Copy;
          WB.WorkSheets[IPage].Columns[ExcelColumnNumToChar(Analiz.FirstColDat+1)+':'+ExcelColumnNumToChar(Col2)].Insert(xlDown, true);
          WB.WorkSheets[IPage].Range[ExcelCombineAddress(Analiz.FirstColDat, Row1-Analiz.IRowBefore), ExcelCombineAddress(Col2, Row1+Analiz.IRowAfter)].Borders[xlInsideVertical].Weight := xlThin;
       end;
       {Размеры таблицы: по вертикали}
       If Analiz.ILoop > 1 then begin
          WB.WorkSheets[IPage].Rows[IntToStr(Row1)].Copy;
          WB.WorkSheets[IPage].Rows[IntToStr(Row1+1)+':'+IntToStr(Row2-Analiz.IRowAfter)].Insert(xlDown, true);
          WB.WorkSheets[IPage].Range[ExcelCombineAddress(Col1, Row1), ExcelCombineAddress(Col2, Row2)].Borders[xlInsideHorizontal].Weight := xlThin;
       end;
       Goto Dl;
    end;

    {===  Отчет: не обнаружен  ================================================}
    ErrMsg('Не определен тип аналитического отчета!'); Exit;

    Dl:
    If Analiz.LRegions.Count=0  then Exit;
    If (Col1 > Col2) or (Row1 > Row2) then Exit;
    S2 := ExcelCombineAddress(Col2, Row2);
    {Записываем область данных}
    S := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_DAT, '');
    If S <> '' then begin
       S3 := ExcelCombineAddress(Analiz.FirstColDat, Row1);
       WB.WorkSheets[IPage].Range[S].Value := S3+':'+S2;
    end;

    {===  Заполняем область данных  ===========================================}
    AData := VarArrayCreate([Row1, Row2, Col1, Col2], varVariant);
    try
       {Просматриваем все столбцы области}
       Analiz.IPeriod := Analiz.IStep * (-1) * (Analiz.ILoop - 1);
       For ICol := Col1 to Col2 do begin
           If Not CreateColumn(ICol) then Exit;
       end;
       Range := WB.WorkSheets[IPage].Range[S1, S2];

       {Для матрицы сбрасываем форматирование}
       If IDMatrix > 0 then begin
          Range.NumberFormat := '@';
          Range.FormatConditions.Delete;
          Range.Validation.Delete; // Ошибка, когда разные условия: в старом 1Е-4 исправил
       end;

       {Записываем область данных в отчет}
       Range.Value := AData;
       Range       := Unassigned;

       {Возвращаемый результат}
       Result:=true;
    finally
       {Освобождаем память}
       If Length(LParamNew) > 0 then begin
          Analiz.PLParam := PParamOld;
          SetLength(LParamNew, 0);
       end;
       VarClear(AData);
       AData := Unassigned;
       FBASE.ClearCash;
    end;
end;




