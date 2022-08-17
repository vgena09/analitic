{==============================================================================}
{====================     Формирует область ввода    ==========================}
{==============================================================================}
{====================   SPage  = 1, 5                ==========================}
{====================   SBlock = 'A5:BH125'          ==========================}
{====================   SBlock = S1+':'+S2           ==========================}
{====================   SBlock = Row1,Col1,Row2,Col2 ==========================}
{==============================================================================}
function TREPORT.CreateBlockStat: Boolean;
var Row1, Col1       : Integer;
    Row2, Col2       : Integer;
    IRow, ICol       : Integer;
    SKey             : String;
    S, S1, S2        : String;
    AData, Range     : Variant;
begin
    {Инициализация}
    Result:=false;
    S1 := AnsiUpperCase(CutSlovo(SBlock, 1, ':'));
    S2 := AnsiUpperCase(CutSlovo(SBlock, 2, ':'));
    If S1='' then Exit;
    If S2='' then S2:=S1;

    {Информатор}
    SetInfo('');

    {Определяем область}
    If Not ExcelSeparatAddress(S1, Col1, Row1) then Exit;
    If Not ExcelSeparatAddress(S2, Col2, Row2) then Exit;
    If (Col1>Col2) or (Row1>Row2) then Exit;
    AData := VarArrayCreate([Row1, Row2, Col1, Col2], varVariant);
    try
       {Просматриваем все ячейки области}
       For ICol := Col1 to Col2 do begin

          IRow := Row1;
          While IRow<=Row2 do begin

             {Переводим адрес ячейки в форму A1}
             S:=ExcelCombineAddress(ICol, IRow);
             If S='' then Exit;

             {Читаем и разбиваем код для текущей ячейки}
             SKey:=Trim(FIni.ReadString(INI_PAGE+SPage, S, ''));
             If SKey='' then begin Inc(IRow); Continue; end;

             {Запоминаем значение}
             If IDMatrix = 0 then AData[IRow, ICol] := ValueCell(SKey, CorrMonth, '')  // CorrMonth при создании нового статотчета с импортом данных за предыдущий период
                             else AData[IRow, ICol] := SKey;
             Inc(IRow);
          end;
       end;
       Range := WB.WorkSheets[IPage].Range[S1, S2];

       {Для матрицы сбрасываем форматирование}
       If IDMatrix > 0 then begin
          Range.NumberFormat := ' @';       // Office > 2003 преобразует '@' в '64'
          Range.FormatConditions.Delete;
          Range.Validation.Delete;          // Ошибка, когда разные условия: в старом 1Е-4 исправил
       end;

       {Записываем область}
       Range.Value := AData;
       Result      := true;
    finally
       {Освобождаем память}
       VarClear(AData);
    end;
end;

