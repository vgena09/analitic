{==============================================================================}
{====================  ФОРМИРУЕМ СТОЛБЕЦ АНАЛИТИКИ   ==========================}
{==============================================================================}
function CreateColumn(const ICol: Integer): Boolean;
label Ex;
var IRegion                : Integer;
    SKey, SReg, S1, S2, S0 : String;
    IRow, I                : Integer;
    BAbs                   : Boolean;
    IVal1, IVal2           : Extended;
begin
    {Инициализация}
    Result:=false;

    {Читаем код столбца}
    S1   := ExcelCombineAddress(ICol, Row1);
    SKey := Trim(FIni.ReadString(INI_PAGE+SPage, S1, ''));
    If (Not IsKod) and (SKey = '') then begin Result:=true; Exit; end;
    If IDMatrix > 0 then begin AData[Row1, ICol]:=SKey; Result:=true; Exit; end;
    SetInfo('столбец: '+ExcelColumnNumToChar(ICol));

    {*** СПЕЦИАЛЬНЫЕ КОМАНДЫ ***}
    {Команда: порядковый номер}
    If CmpStr(SKey, 'НОМЕР') then begin
       IRow := Row1;
       I    := 1;
       While IRow <= Row2 do begin
          AData[IRow, ICol]:=I;
          Inc(I);
          If CmpStr(SType, INI_PAGE_TYPE_DYNAMIC) then Inc(IRow) else IRow:=IRow+Analiz.ILoop;
       end;
       Goto Ex;
    end;

    {Команда: регион}
    If CmpStr(SKey, 'РЕГИОН') then begin
       IRow := Row1;
       For I:=0 to Analiz.LRegions.Count-1 do begin
          AData[IRow, ICol]:=Analiz.LRegions[I];
          IRow:=IRow+Analiz.ILoop;
       end;
       If CmpStrList(SType, [INI_PAGE_TYPE_STANDART, INI_PAGE_TYPE_COMPARE]) >=0 then
          AData[IRow, ICol]:='В абс. цифрах'+CHR(10)+'рост+, сниж.-';
       Goto Ex;
    end;

    {Команда: период}
    If FindStr('ПЕРИОД', SKey)=1 then begin
       SKey:=CutModulChar(SKey, '(', ')');
       If CmpStr(SKey, 'ГГГГ') then begin   // Не доделано: при Analiz.ILoop>2
          //IncDate(EncodeDate(StrToInt(SYear), SMonth, 1), 0, 0, -1);
          S1:=IntToStr(StrToInt(SYear)-1);
          S2:=SYear;
          IRow:=Row1;
          While IRow < Row2 do begin
             AData[IRow,   ICol] := S1;
             AData[IRow+1, ICol] := S2;
             IRow:=IRow+Analiz.ILoop;
          end;
          Goto Ex;
       end;
       SKey := 'ПЕРИОД';
    end;

    {**************************************************************************}
    {*** Произвольная аналитика: рейтинг **************************************}
    {**************************************************************************}
    If CmpStr(SType, INI_PAGE_TYPE_RATING) then begin
       {Команда: код}
       If ICol >= Analiz.FirstColDat then begin
          IsKod := true;
          {Индекс строки параметров}
          I := ICol - Analiz.FirstColDat;
          If (I < Low(Analiz.PLParam^)) or (I > High(Analiz.PLParam^)) then Goto Ex;
          {Заголовок}
          S1 := ExcelCombineAddress(ICol, FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_CAPTION, 45));
          WB.WorkSheets[IPage].Range[S1, S1].Value := Analiz.PLParam^[I].SCaption;
          {Формула}
          SKey := Analiz.PLParam^[I].SFormula;
          {Формат и условный формат}
          S1 := ExcelCombineAddress(ICol, Row1);
          S2 := ExcelCombineAddress(ICol, Row2);
          If IsFormulaDiv(SKey)
          then WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_FLOAT, '# ##0,0;-# ##0,0;;@')
          else WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_INT,   '# ###;-# ###;;@');
          {Цикл по регионам}
          IRow:=Row1;
          For IRegion:=0 to Analiz.LRegions.Count-1 do begin
             AData[IRow, ICol]:=ValueCell(SKey, 0, Analiz.LRegions[IRegion]);
             Inc(IRow);
          end;
       end;
       Goto Ex;
    end;

    {**************************************************************************}
    {*** Произвольная аналитика: регион ***************************************}
    {**************************************************************************}
    If CmpStr(SType, INI_PAGE_TYPE_REGION) then begin
       {Команда: заголовок}
       If CmpStr(SKey, 'ЗАГОЛОВОК') then begin
          For I:=0 to Length(Analiz.PLParam^)-1 do begin
             AData[Row1+I*Analiz.ILoop, ICol]:=Analiz.PLParam^[I].SCaption;
          end;
          Goto Ex;
       end;
       {Команда: цвет}
       If CmpStr(SKey, 'ЦВЕТ') then begin
          For I:=0 to Length(Analiz.PLParam^)-1 do begin
             AData[Row1+I*Analiz.ILoop+1, ICol]:=Analiz.PLParam^[I].SColor;
          end;
          Goto Ex;
       end;
       {Команда: код}
       If ICol >= Analiz.FirstColDat then begin
          IsKod := true;
          {Название региона}
          I    := ICol - Analiz.FirstColDat;
          If I > Analiz.LRegions.Count - 1 then Goto Ex;
          SReg := Analiz.LRegions[I];
          {Заголовок}
          S1 := ExcelCombineAddress(ICol, FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_REGION, 46));
          WB.WorkSheets[IPage].Range[S1, S1].Value := SReg;
          {Заполняем столбец}
          IRow:=Row1;
          For IRegion:=0 to Length(Analiz.PLParam^)-1 do begin
             SKey := Analiz.PLParam^[IRegion].SFormula;
             {Формат}
             If ICol = Analiz.FirstColDat then begin
                S1 := ExcelCombineAddress(Analiz.FirstColDat, IRow);
                S2 := ExcelCombineAddress(Col2, IRow-1+Analiz.ILoop);
                If IsFormulaDiv(SKey)
                then WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_FLOAT, '# ##0,0;-# ##0,0;;@')
                else WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_INT,   '# ###;-# ###;;@');
             end;
             {Значения}
             For I:=Analiz.ILoop-1 downto 0 do begin
                AData[IRow, ICol]:=ValueCell(SKey, I*Analiz.IStep*(-1), SReg);
                Inc(IRow);
             end;
          end;
       end;
       Goto Ex;
    end;


    {**************************************************************************}
    {*** Произвольная аналитика: динамика *************************************}
    {**************************************************************************}
    If CmpStr(SType, INI_PAGE_TYPE_DYNAMIC) then begin
       {Команда: заголовок}
       If CmpStr(SKey, 'ПЕРИОД') then begin
          For I:=0 to Analiz.ILoop-1 do begin
             AData[Row2-I, ICol] := CorrectPeriod(SYear, SMonth, I * (-1)); // DateToStr(Dat), System.SysUtils.FormatDateTime('yyyy mmmm', Dat);
          end;
          Goto Ex;
       end;
       {Команда: код}
       If ICol >= Analiz.FirstColDat then begin
          IsKod := true;
          {Индекс строки параметров}
          I := ICol - Analiz.FirstColDat;
          If (I < Low(Analiz.PLParam^)) or (I > High(Analiz.PLParam^)) then Goto Ex;
          {Раскраска}
          S1 := ExcelCombineAddress(ICol, FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_COLOR, 41));
          WB.WorkSheets[IPage].Range[S1, S1].Value := Analiz.PLParam^[I].SColor;
          {Заголовок}
          S1 := ExcelCombineAddress(ICol, FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_CAPTION, 45));
          WB.WorkSheets[IPage].Range[S1, S1].Value := Analiz.PLParam^[I].SCaption;
          {Формула}
          SKey := Analiz.PLParam^[I].SFormula;
          {Формат и условный формат}
          S1 := ExcelCombineAddress(ICol, Row1);
          S2 := ExcelCombineAddress(ICol, Row2);
          If IsFormulaDiv(SKey)
          then WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_FLOAT, '# ##0,0;-# ##0,0;;@')
          else WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_INT,   '# ###;-# ###;;@');
          {Цикл по периодам}
          IRow:=Row1;
          For I:=Analiz.ILoop-1 downto 0 do begin
             If IsCtrl then begin
                AData[IRow, ICol] := ValueCell(SKey, I * Analiz.IStep * (-1), SRegion);
             end else begin
                If FormatDateTime('m', CorrectPeriod(SYear, SMonth, I * (-1))) = '1'
                then AData[IRow, ICol] := ValueCell(SKey, I * Analiz.IStep * (-1), SRegion)
                else AData[IRow, ICol] := ValueCell(SKey, I * Analiz.IStep * (-1), SRegion) - ValueCell(SKey, (I+1) * Analiz.IStep * (-1), SRegion);
             end;
             Inc(IRow);
          end;
       end;
       Goto Ex;
    end;


    {**************************************************************************}
    {*** Произвольная аналитика: сравнение ************************************}
    {**************************************************************************}
    If CmpStr(SType, INI_PAGE_TYPE_COMPARE) then begin
       {Команда: код}
       If ICol >= Analiz.FirstColDat then begin
          IsKod := true;
          {Индекс строки параметров}
          I := ICol - Analiz.FirstColDat;
          If (I < Low(Analiz.PLParam^)) or (I > High(Analiz.PLParam^)) then Goto Ex;
          {Раскраска}
          S1 := ExcelCombineAddress(ICol, FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_COLOR, 41));
          WB.WorkSheets[IPage].Range[S1, S1].Value := Analiz.PLParam^[I].SColor;
          {Заголовок}
          S1 := ExcelCombineAddress(ICol, FIni.ReadInteger(INI_PAGE+SPage, INI_PAGE_ROW_CAPTION, 45));
          WB.WorkSheets[IPage].Range[S1, S1].Value := Analiz.PLParam^[I].SCaption;
          {Формула}
          SKey := Analiz.PLParam^[I].SFormula;
          {Формат и условный формат}
          S1 := ExcelCombineAddress(ICol, Row1);
          S2 := ExcelCombineAddress(ICol, Row2);
          If IsFormulaDiv(SKey)
          then WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_FLOAT, '# ##0,0;-# ##0,0;;@')
          else WB.WorkSheets[IPage].Range[S1, S2].NumberFormat := FIni.ReadString(INI_PAGE+SPage, INI_FORMAT_INT,   '# ###;-# ###;;@');
       end;
    end;

    {Признак +/-}
    BAbs:=(SKey[1]='!');
    If BAbs then begin
       Delete(SKey, 1, 1);
       SKey:=Trim(SKey);
    end;

    {Альтернативный код}
    S1   := '';
    S2   := SKey;
    SKey := Trim(TokChar(S2, '|'));
    S2   := Trim(S2);
    If S2<>'' then begin
       S1 := Trim(TokChar(S2, ':'));
       S2 := Trim(S2);
    end;
    S0   := SKey;

    {Цикл по периодам и регионам}
    IRow:=Row1;
    For IRegion:=0 to Analiz.LRegions.Count-1 do begin
       For I:=Analiz.ILoop-1 downto 0 do begin
          If CmpStr(S1, Analiz.LRegions[IRegion]) then SKey:=S2 else SKey:=S0;
          AData[IRow, ICol]:=ValueCell(SKey, I*Analiz.IStep*(-1), Analiz.LRegions[IRegion]);
          Inc(IRow);
       end;
    end;

    {Динамика в абсолютных значениях}
    If Not BAbs then begin
       For I:=Analiz.ILoop-1 downto 0 do begin
          IVal1 := ValueCell(SKey, I*Analiz.IStep*(-1),     Analiz.LRegions[Analiz.LRegions.Count-1]);
          IVal2 := ValueCell(SKey, (I+1)*Analiz.IStep*(-1), Analiz.LRegions[Analiz.LRegions.Count-1]);
          AData[IRow, ICol]:=IVal1 - IVal2;
          Inc(IRow);
       end;
    end;

Ex: {Возвращаемый результат}
    Result:=true;
end;


