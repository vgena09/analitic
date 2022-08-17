unit FunText;

interface

uses System.SysUtils, System.Classes, System.Math,
     Vcl.Dialogs, IdGlobal, RxParsing,
     FunType;

{Количество символов Ch в строке S}
function  GetColChar(const S: String; const Ch: Char): Integer;
{Количество строк S0 в строке S}
function  GetColStr(const S, S0: String): Integer;
{Количество слов в строке Str, STerm - разделитель слов}
function  GetColSlov(const Str: String; const STerm: String): Integer;
{Вырезать со строки Str слово ISlovo, STerm - разделитель слов}
function  CutSlovo(const Str: String; const ISlovo: Integer; const STerm: String): String;
{Вырезать со строки S слово I с конца, Ch - разделитель слов}
function  CutSlovoEndChar(const S: String; const I: Integer; const Ch: Char): String;
{Заменить в строке S слово I с конца на S0, Ch - разделитель слов}
function  ReplSlovoEndChar(const S, S0: String; const I: Integer; const Ch: Char): String;
{Вырезает со строки Str ILen символов начиная с позиции IPos}
function  CutStrPos(const Str: String; const IPos, ILen: Integer): String;
{Обрезает в строке Str свыше ILen символов, заменяя их на ...}
function  CutLongStr(const Str: String; const ILen: Integer): String;
{Заменяет в строке Dest ILen символов начиная с позиции IPos на строку SubStr}
function  ReplStrPos(const Dest: String; const IPos, ILen: Integer; const SubStr: String): String;
{Заменяет в строке S все подстроки S1 на S2}
function  ReplStr(const S: String; const S1, S2: String): String;
{Заменяет в строке S все подстроки S1 на S2 (циклическое выполнение)}
function  ReplStrLoop(const S: String; const S1, S2: String): String;
{Проверить содержит ли строка S строку SFind - нужно использовать не эту функцию, а функцию Pos}
function  InStrMy(const StartPos:Integer; const S, SFind: String): Integer;
{Вырезаем блок со строки SBlock, ограниченный STerm и обрезаем пробелы}
function  CutBlock(var SBlock: String; const STerm: String): String;
{Вырезаем блок со строки, ограниченный символами С}
function  TokChar(var Si: String; const C: Char): String;
{Вырезаем блок со строки, ограниченный строкой S}
function  TokStr(var Si: String; const S: String): String;
{Вырезаем блок с конца строки, ограниченный символами С}
function  TokCharEnd(var Si: String; const C: Char): String;

{Вырезает из строки S1.S2. ... строку S1}
function  CutClass(var Si: String): String;

{Определяет позицию первого символа первого вложенного блока, ограниченного символами С1-С2}
function  GetPosStartChar(const S: String; const C1, C2: Char): Integer;
{Определяет позицию первого символа первого вложенного блока, ограниченного строками S1-S2}
function  GetPosStartStr(const S: String; const S1, S2: String): Integer;
{Определяет позицию первого символа первого пакета SPackType}
function  GetPosStartPack(const S, SPackType: String): Integer;
{Определяет позицию последнего символа первого вложенного блока, ограниченного символами С1-С2}
function  GetPosEndChar(const S: String; const StartPos: Integer; const C2:Char): Integer;
{Определяет позицию последнего символа первого вложенного блока, ограниченного строками S1-S2}
function  GetPosEndStr(const S: String; const StartPos: Integer; const S2: String): Integer;
{Определяет позицию последнего символа первого пакета SPackType}
function  GetPosEndPack(const S, SPackType: String): Integer;

{Объединяет два массива в один}
function  SumArrayStr(const S1, S2: array of String): TArrayStr;
{Объединяет два StringList в один, отбирая только уникальные записи}
procedure SumSListUniq(const PList, PListAdd: PStringList);
{Разделяет мультистроку в массив}
function  SeparatMStr(const MStr: String; const PLStr: PStringList; const Separator: String): Boolean;
{Собирает массив в мультистроку}
function  GroupMStr(const PLStr: PStringList; const Separator: String): String;

{Вырезает первый вложенный блок, ограниченный символами С1 и С2}
function  CutModulChar(const S: String; const C1, C2: Char): String;
{Вырезает первый вложенный блок, ограниченный строками S1 и S2}
function  CutModulStr(const S: String; const S1, S2: String): String;

{Вырезает первый пакет SPack из списка}
function  CutModulPackList(const Str           : String;
                           const SPackTypeList : array of String;
                           var   SPack         : String): Integer;
{Вырезает первый пакет типа SPackType}
function  CutModulPack(const Str, SPackType: String): String;
{Заменяет первый вложенный блок, ограниченный символами С1 и С2}
function  ReplModulChar(const Str1, Str2: String; const C1, C2: Char): String;
{Заменяет первый вложенный блок, ограниченный строками S1 и S2}
function  ReplModulStr(const Str1, Str2: String; const S1, S2: String): String;
{Заменяет первый пакет типа SPackType}
function  ReplModulPack(const Str1, Str2, SPackType: String): String;
{Сравнивает строку с перечнем строк при разных регистрах}
function  CmpStrList(const S: String; SList: array of String): Integer;
{Сравнивает 2 строки при разных регистрах}
function  CmpStr(const S1, S2: String): Boolean;
{Число первых одинаковых символов при разных регистрах}
function  CmpStrFirst(const S1, S2: String): Integer;
{Ищет субстроку SubStr в строке Str без учета регистра}
function  FindStr(const SubStr, Str: String): Integer;
{Содержит ли формула операцию деления}
function  IsFormulaDiv(const SFormula: String): Boolean;


{Вырезает ILine (0, 1, ...) строку мультистроки}
function  CutLine(const S: String; const ILine: Integer): String;
{Обрезает в конце строки символы Сhr(10) и Chr(13)}
function  CutEndStr(const S: String): String;
{Преобразует в строке Str: начиная с позиции IPos заглавные буквы в кол-ве ICount}
function  UpperStrPart(const Str: String; const IPos, ICount: Integer): String;
{Преобразует массив строк в заглавные}
function  UpperArrayStr(const AStr: TArrayStr): TArrayStr;
{Поиск заданной строки в массиве строк}
function  FindStrInArray(const StrFind: String; const StrArray: array of String): Boolean;
{Поиск заданного числа в массиве чисел}
function  FindIntInArray(const IntFind: Integer; const IntArray: array of Integer): Boolean;

{Является ли строка целым числом}
function  IsIntegerStr(const S: String): Boolean;
{Является ли строка дробным числом}
function  IsFloatStr(const S: String): Boolean;
{Вычисление вражения с обработкой деления на ноль}
function  CalcStr(const SFormula: String): Extended;

{Вырезает значения секции Section из списка PSrc^ в список PDest^}
function  CutLSection(const PSrc, PDest: PStringList; const Section: String): Boolean;
{Удаляет из списка PLst^ секцию Section}
function  DelLSection(const PLst: PStringList; const Section: String): Boolean;
{Оставляет в списке PLst^ только элементы ElName.*}
function  CutLIDE(const PLst: PStringList; const ElName: String): Boolean;
{Вырезает значения элемента из списка в новый список}
function  CutLKey(const PSrc, PDest: PStringList; const ElName: String): Boolean;


{Вырезает из пути имя файла без расширения}
function  ExtractFileNameWithoutExt(const FPath: String): String;


{Замена в строке Str блоков, ограниченных символами Ch1 и Ch2, на псевдооператоры: |%N|%}
function  ReplPsevdoModul(const Str: String; var PsStr: TStringList; const Ch1, Ch2: Char): String;
{Количество слов в псевдостроке Str, Str0 - разделитель слов}
function  GetColPSlovStr(const Str: String; const Str0: String): Integer;
{Вырезать с псевдостроки Str слово I, Str0 - разделитель слов}
function  CutPSlovoStr(const Str: String; const I: Integer; const Str0: String): String;

{Является ли строка кодом}
function  IsKod(const SKod: String): Boolean;
{Разбивает код на составляющие}
function  SeparatKod(const SKod: String; var STab, SRow, SCol: String): Boolean;
{Формирует код из составляющих}
function  GeneratKod(const STab, SRow, SCol: String): String;

implementation

uses FunConst;

{$INCLUDE FunText_Modul}


{==============================================================================}
{========================  О Б Щ И Е   Ф У Н К Ц И И  =========================}
{==============================================================================}


{==============================================================================}
{===================   Количество символов Ch в строке S  =====================}
{==============================================================================}
function GetColChar(const S: String; const Ch: Char): Integer;
var I: Integer;
begin
    Result:=0;
    If S='' then Exit;                        
    For I:=1 to Length(S) do
       If S[I]=Ch then Result:=Result+1;
end;


{==============================================================================}
{====================   Количество строк S0 в строке S  =======================}
{==============================================================================}
function GetColStr(const S, S0: String): Integer;
label Nx;
var I: Integer;
begin
    Result:=0;
    If S='' then Exit;
    I:=0;
Nx: I:=InStrMy(I, S, S0); If I=0 then Exit;
    Result:=Result+1;
    I:=I+Length(S0);
    goto Nx;
end;


{==============================================================================}
{=========  Количество слов в строке Str, STerm - разделитель слов  ===========}
{==============================================================================}
function GetColSlov(const Str: String; const STerm: String): Integer;
begin
    Result:=0;
    While CutSlovo(Str, Result+1, STerm)<>'' do Result:=Result+1;
end;


{==============================================================================}
{======  Вырезать со строки Str слово ISlovo, STerm - разделитель слов  =======}
{==============================================================================}
function CutSlovo(const Str: String; const ISlovo: Integer; const STerm: String): String;
var S : String;
    I : Integer;
begin
    Result := '';
    S      := Str;
    For I := 1 to ISlovo do Result := CutBlock(S, STerm);
end;


{==============================================================================}
{======= Вырезать со строки S слово I c конца, Ch - разделитель слов ==========}
{==============================================================================}
function CutSlovoEndChar(const S: String; const I: Integer; const Ch: Char): String;
label Nx;
var   J, Slov : Integer;
begin
    Result:='';
    if Length(S)=0 then Exit;
    J:=Length(S); Slov:=0;

Nx: While (S[J]=Ch) and (J>=1) do J:=J-1;
    If J=0 then Exit;
    Slov:=Slov+1;
    While (S[J]<>Ch) and (J>=1) do begin
        If Slov=I then Result:=S[J]+Result;
        J:=J-1;
    end;
    If J=0 then Exit;
    If Result='' then goto Nx;
end;


{==============================================================================}
{====== Заменить в строке S слово I с конца на S0, Ch - разделитель слов ======}
{==============================================================================}
{======   Для упрощения функции не обрабатывается первое с начала слово  ======}
{==============================================================================}
function ReplSlovoEndChar(const S, S0: String; const I: Integer; const Ch: Char): String;
var I0, ILen, ISlovo : Integer;
begin
    {Инициализация}
    Result := S;
    ISlovo := 0;
    ILen   := 0;
    If Length(Result)=0 then Exit;

    {Ищем нужное слово}
    For I0:=Length(Result) downto 1 do begin
       If Result[I0]=Ch then begin
          ISlovo := ISlovo+1;
          {Слово найдено}
          If ISlovo=I then begin
             Delete(Result, I0+1, ILen);
             Insert(S0, Result, I0+1);
             Break;
          end;
          ILen := 0;
       end else begin
          ILen := ILen+1;
       end;
    end;
end;


{==============================================================================}
{======== Вырезает со строки Str ILen символов начиная с позиции IPos =========}
{==============================================================================}
function CutStrPos(const Str: String; const IPos, ILen: Integer): String;
var I: Integer;
begin
    Result:='';
    If (IPos+ILen-1)>Length(Str) then Exit;
    For I:=IPos to IPos+ILen-1 do Result:=Result+Str[I];
end;


{==============================================================================}
{========  Обрезает в строке Str свыше ILen символов, заменяя их на ...  ======}
{==============================================================================}
function CutLongStr(const Str: String; const ILen: Integer): String;
begin
    If Length(Str)>ILen then Result:=CutStrPos(Str, 1, ILen-3)+'...'
                        else Result:=Str;
end;


{==============================================================================}
{=Заменяет в строке Dest ILen символов начиная с позиции IPos на строку SubStr=}
{==============================================================================}
function ReplStrPos(const Dest: String; const IPos, ILen: Integer; const SubStr: String): String;
begin
    Result:=Dest;
    Delete(Result, IPos, ILen);
    Insert(SubStr, Result, IPos);
end;


{==============================================================================}
{=============  Заменяет в строке S все подстроки S1 на S2  ===================}
{==============================================================================}
function ReplStr(const S: String; const S1, S2: String): String;
const MASKREPL = '%|!A%|U!%|!T%|O!%|!D%|O!%|!C%|';
var I: Integer;

    function ReplStrSimpl(const S: String; const S1, S2: String): String;
    begin
        {Инициализация}
        Result:=S;
        I:=Pos(S1, Result);

        While I>0 do begin
           Result:=ReplStrPos(Result, I, Length(S1), S2);
           I:=Pos(S1, Result);
        end;
    end;

begin
    Result:=ReplStrSimpl(S,      S1,       MASKREPL);
    Result:=ReplStrSimpl(Result, MASKREPL, S2);
end;


{==============================================================================}
{===== Заменяет в строке S все подстроки S1 на S2 (циклическое выполнение) ====}
{==============================================================================}
function ReplStrLoop(const S: String; const S1, S2: String): String;
var S0: String;
begin
    {Инициализация}
    Result := '';
    S0     := S;

    While S0<>Result do begin
       Result := S0;
       S0     := ReplStr(Result, S1, S2);
    end;
end;


{==============================================================================}
{=============   Проверить содержит ли строка S строку SFind    ===============}
{============= Нужно использовать не эту функцию, а функцию Pos ===============}
{==============================================================================}
function InStrMy(const StartPos:Integer; const S, SFind: String): Integer;
var I, LengthFind: Integer;
begin
     Result     := 0;
     LengthFind := Length(SFind);
     If (S='') or (SFind='') or (Length(S)<LengthFind) then Exit;
     for I:=StartPos to Length(S)-LengthFind+1 do
         If String(Copy(S, I, LengthFind)) = SFind then begin Result:=I; break; end;
end;


{==============================================================================}
{==== Вырезаем блок со строки SBlock, ограниченный STerm и обрезаем пробелы ===}
{==============================================================================}
function CutBlock(var SBlock: String; const STerm: String): String;
var I, ILen:  Integer;
begin
    ILen := Length(STerm);
    I := Pos(STerm, SBlock);
    If I = 0 then begin
       Result := Trim(SBlock);
       SBlock := '';
    end else begin
       Result := Trim(Copy(SBlock, 1, I-1));
       Delete(SBlock, 1, I+ILen-1);
       SBlock := Trim(SBlock);
    end;
end;


{==============================================================================}
{============ Вырезаем блок со строки, ограниченный символами С ===============}
{==============================================================================}
function TokChar(var Si: String; const C: Char): String;
var I:  Integer;
begin
    Result:=''; if Si='' then Exit;
    I:=1;
    while (Si[I]<>C)and(I<=Length(Si)) do begin
        Result:=Result+Si[I]; I:=I+1;
    end;
    Delete(Si,1,I);
end;

{==============================================================================}
{============ Вырезаем блок со строки, ограниченный строкой S =================}
{==============================================================================}
function TokStr(var Si: String; const S: String): String;
var I:  Integer;
begin
    Result:=Si;
    if (Si<>'') and (S<>'') then I:=InStrMy(1, Si, S) else I:=0;
    If I>0 then begin
       Result:=Copy(Si, 1, I-1);
       Delete(Si,1,I+Length(S)-1);
    end else begin
       Si:='';
    end;
end;

{==============================================================================}
{=========  Вырезаем блок с конца строки, ограниченный символами С  ===========}
{==============================================================================}
function TokCharEnd(var Si: String; const C: Char): String;
var I:  Integer;
begin
    Result:=''; if Si='' then Exit;
    I:=Length(Si);
    while (Si[I]<>C)and(I>=1) do begin
        Result:=Si[I]+Result; I:=I-1;
    end;
    {Удаляем вырезанный блок}
    Delete(Si, Length(Si)-Length(Result)+1, Length(Result));
    {Удаляем символ-разделитель}
    Delete(Si, Length(Si), 1);
end;


{==============================================================================}
{===================  Вырезает из строки S1.S2. ... строку S1  ================}
{==============================================================================}
function CutClass(var Si: String): String;
var I: Integer;
begin
    Si:=Trim(Si);
    Result:=CutPSlovoStr(Si, 1, '.');
    Delete(Si, 1, Length(Result));
    I:=1;
    While I<=Length(Si) do begin
       If Si[I]='.' then begin
          Delete(Si, 1, 1);
          Break;
       end;
       Delete(Si, 1, 1);
       Inc(I);
    end;
end;


{==============================================================================}
{======================  Объединяет два массива в один  =======================}
{==============================================================================}
function SumArrayStr(const S1, S2: array of String): TArrayStr;
var I, L1, L2: Integer;
begin
    L1:=Length(S1);
    L2:=Length(S2);
    SetLength(Result, L1+L2);
    For I:=0 to L1-1 do Result[I]    := S1[I];
    For I:=0 to L2-1 do Result[L1+I] := S2[I];
end;



{==============================================================================}
{===== Объединяет два StringList в один, отбирая только уникальные записи =====}
{==============================================================================}
procedure SumSListUniq(const PList, PListAdd: PStringList);
var I: Integer;
begin
    For I:=0 to PListAdd^.Count-1 do begin
       If PList^.IndexOf(PListAdd^[I])<0 then PList^.Add(PListAdd^[I]);
    end;
end;


{==============================================================================}
{=====================  Разделяет мультистроку в массив  ======================}
{==============================================================================}
function SeparatMStr(const MStr: String; const PLStr: PStringList; const Separator: String): Boolean;
var S: String;
    I: Integer;
begin
    {Инициализация}
    Result:=false;
    If PLStr=nil then Exit;
    PLStr^.Clear;

    {Просматриваем все статусы и формируем список}
    I:=1; S:=CutSlovo(MStr, I, Separator);
    While S<>'' do begin
       Result:=true;
       PLStr^.Add(S);
       Inc(I); S:=CutSlovo(MStr, I, Separator);
    end;
end;


{==============================================================================}
{=====================  Собирает массив в мультистроку   ======================}
{==============================================================================}
function GroupMStr(const PLStr: PStringList; const Separator: String): String;
var I: Integer;
begin
    {Инициализация}
    Result:='';
    If PLStr=nil then Exit;

    {Просматриваем все статусы и формируем строку}
    For I:=0 to PLStr^.Count-1 do Result:=Result+Separator+PLStr^[I];

    {Удаляем последние символа разделителя}
    If Result<>'' then Delete(Result, 1, Length(Separator));
end;


{==============================================================================}
{========  Сравнивает строку с перечнем строк при разных регистрах  ===========}
{==============================================================================}
function CmpStrList(const S: String; SList: array of String): Integer;
var I  : Integer;
    S0 : String;
begin
    {Инициализация}
    Result:=-1;
    S0:=Trim(AnsiUpperCase(S));

    For I:=Low(SList) to High(SList) do begin
       If S0=Trim(AnsiUpperCase(SList[I])) then begin
          Result:=I;
          Break;
       end;
    end;
end;


{==============================================================================}
{==============  Cравнивает 2 строки при разных регистрах  ====================}
{==============================================================================}
function CmpStr(const S1, S2: String): Boolean;
begin
    Result:=(AnsiUpperCase(S1)=AnsiUpperCase(S2));
end;


{==============================================================================}
{=========   Число первых одинаковых символов при разных регистрах   ==========}
{==============================================================================}
function CmpStrFirst(const S1, S2: String): Integer;
var I: Integer;
begin
    {Инициализация}
    Result:=0;

    For I:=1 to Min(Length(S1), Length(S2)) do begin
       If CmpStr(S1[I], S2[I]) then Result:=I
                               else Break;
    end;
end;



{==============================================================================}
{=================  Содержит ли формула операцию деления  =====================}
{==============================================================================}
function IsFormulaDiv(const SFormula: String): Boolean;
var S, S0 : String;
begin
    {Инициализация}
    Result := false;
    S      := SFormula;
    S0     := CutModulChar(S, '[', ']');
    {Если формула}
    If S0 <> '' then begin
       While S0 <> '' do begin
          S   := ReplModulChar(S, '', '[', ']');
          S0  := CutModulChar(S, '[', ']');
       end;
       Result := Pos('/', S) > 0;
    end;
end;


{==============================================================================}
{==========  Ищет субстроку SubStr в строке Str без учета регистра  ===========}
{==============================================================================}
function FindStr(const SubStr, Str: String): Integer;
begin
    Result:=Pos(AnsiUpperCase(SubStr), AnsiUpperCase(Str));
end;



{==============================================================================}
{============  Вырезает ILine (0, 1, ...) строку мультистроки  ================}
{==============================================================================}
function CutLine(const S: String; const ILine: Integer): String;
var SList: TStringList;
begin
    {Инициализация}
    Result:='';
    SList:=TStringList.Create;

    try
       SList.Text:=S;
       If (SList.Count<0) or (ILine>=SList.Count) then Exit;
       Result:=SList[ILine];
    finally
       SList.Free;
    end;
end;


{==============================================================================}
{==========  Обрезает в конце строки символы Сhr(10) и Chr(13)  ===============}
{==============================================================================}
function CutEndStr(const S: String): String;
begin
    {Инициализация}
    Result:=S;

    While true do begin
       If Length(Result)=0 then Break;
       If (Result[Length(Result)]<>Chr(10)) and (Result[Length(Result)]<>Chr(13)) then Break;
       Delete(Result, Length(Result), 1);
    end;
end;


{==============================================================================}
{Преобразует в строке Str: начиная с позиции IPos заглавные буквы в кол-ве ICount}
{==============================================================================}
function UpperStrPart(const Str: String; const IPos, ICount: Integer): String;
var S: String;
    I: Integer;
begin
    Result:=Str;
    If Str='' then Exit;
    If (IPos+ICount) > Length(Str) then Exit;
    for I:=IPos to IPos+ICount-1 do begin
        S:=AnsiUpperCase(Str[I]);
        Result[I]:=S[1];
    end;
end;


{==============================================================================}
{==================  Преобразует массив строк в заглавные  ====================}
{==============================================================================}
function UpperArrayStr(const AStr: TArrayStr): TArrayStr;
var I: Integer;
begin
    SetLength(Result, Length(AStr));
    For I:=Low(AStr) to High(AStr) do Result[I]:=AnsiUpperCase(AStr[I]);
end;


{==============================================================================}
{==================  Поиск заданной строки в массиве строк  ===================}
{==============================================================================}
function FindStrInArray(const StrFind: String; const StrArray: array of String): Boolean;
var StrFind0 : String;
    I        : Integer;
begin
    {Инициализация}
    Result   := false;
    StrFind0 := AnsiUpperCase(StrFind);

    {Просматриваем и сравниваем элементы массива}
    For I:=Low(StrArray) to High(StrArray) do begin
       If AnsiUpperCase(StrArray[I])=StrFind0 then begin
          Result:=true;
          Break;
       end;
    end;
end;


{==============================================================================}
{==================  Поиск заданного числа в массиве чисел  ===================}
{==============================================================================}
function FindIntInArray(const IntFind: Integer; const IntArray: array of Integer): Boolean;
var I: Integer;
begin
    Result:=false;
    For I:=Low(IntArray) to High(IntArray) do begin
       If IntArray[I]=IntFind then begin
          Result:=true;
          Break;
       end;
    end;
end;


{==============================================================================}
{==================  Является ли строка целым числом  =========================}
{==============================================================================}
function IsIntegerStr(const S: String): Boolean;
var I: Integer;
begin
    Result:=true;
    If S='' then begin
       Result:=false;
       Exit;
    end;
    For I:=1 to Length(S) do begin
       If IsNumeric(S[I])      then Continue;
       If (I=1) and (S[1]='-') then Continue;
       Result:=false;
       Break;
    end;
end;


{==============================================================================}
{==================  Является ли строка дробным числом  =======================}
{==============================================================================}
function IsFloatStr(const S: String): Boolean;
var I: Integer;
begin
    Result:=true;
    If S='' then begin
       Result:=false;
       Exit;
    end;
    For I:=1 to Length(S) do begin
       If IsNumeric(S[I]) then Continue;
       If ((I=1) and (S[1]='-')) or (S[I]=',') then Continue;
       Result:=false;
       Break;
    end;
end;


{==============================================================================}
{============  Вычисление вражения с обработкой деления на ноль  ==============}
{==============================================================================}
{============                В случае ошибки - 0                 ==============}
{==============================================================================}
function CalcStr(const SFormula: String): Extended;
var SVal, SBlock: String;
begin
    {Инициализация}
    Result := 0;
    SVal   := ReplStrLoop(SFormula, ' ', '');

    {Просматриваем блоки в круглых скобках}
    SBlock := CutModulChar(SVal, '(', ')');
    While SBlock<>'' do begin
       If Pos('/0', SBlock)>0 then Exit;
       SBlock := FloatToStr(GetFormulaValue(SBlock));
       SVal   := ReplModulChar(SVal, SBlock, '(', ')');
       SBlock := CutModulChar(SVal,          '(', ')');
    end;

    {Окончательное вычисление}
    If Pos('/0', SVal)>0 then Exit;
    try    Result := RoundTo(GetFormulaValue(SVal), -4);       // Точность вычислений - цифр после запятой
    except MessageDlg('Ошибка вычисления выражения:'+CH_NEW+SVal+MSGQWEST, mtError, [mbOk], 0);
    end;

    //Result:=ReplStrLoop(Result, ',', '.');
end;


{==============================================================================}
{=====  Вырезает значения секции Section из списка PSrc^ в список PDest^  =====}
{==============================================================================}
{============  [Секция 1]             =========================================}
{============  ...           --> ...  =========================================}
{============  [Секция 2]             =========================================}
{============  ...                    =========================================}
{==============================================================================}
function CutLSection(const PSrc, PDest: PStringList; const Section: String): Boolean;
var I     : Integer;
    B     : Boolean;
    S, S0 : String;
begin
    {Инициализация}
    Result:=false;
    If (PSrc=nil) or (PDest=nil) then Exit;
    PDest^.Clear;
    S0:=AnsiUpperCase(Section);
    B:=false;

    For I:=0 to PSrc^.Count-1 do begin
       S:=Trim(CutModulChar(Trim(PSrc^[I]), '[', ']'));
       If S<>'' then begin
          If AnsiUpperCase(S)=S0 then B:=true else B:=false;
       end else begin
          If B=true then PDest^.Add(PSrc^[I]);
       end;
    end;
    Result:=true;
end;


{==============================================================================}
{============  Удаляет из списка PLst^ секцию Section  ========================}
{==============================================================================}
{============  [Секция 1]        X    =========================================}
{============  ...               X    =========================================}
{============  [Секция 2]             =========================================}
{============  ...                    =========================================}
{==============================================================================}
function DelLSection(const PLst: PStringList; const Section: String): Boolean;
label Nx;
var I     : Integer;
    B     : Boolean;
    S, S0 : String;
begin
    {Инициализация}
    Result := false;
    If PLst=nil then Exit;
    S0     := AnsiUpperCase(Section);
    B      := false;
    I      := 0;

Nx: If I<=PLst^.Count-1 then begin
       {Проверяем название секции}
       S:=Trim(CutModulChar(Trim(PLst^[I]), '[', ']'));
       If S<>'' then begin
          If AnsiUpperCase(S)=S0 then B:=true else B:=false;
       end;
       {Удаляем запись или переходим к следующей}
       If B=true then PLst^.Delete(I)
                 else I:=I+1;
       {Следующая запись}
       Goto Nx;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{===========   Оставляет в списке PLst^ только элементы ElName.*  =============}
{==============================================================================}
{============  Элемент1.Значение1 <-- ElName=Элемент1             =============}
{============  Элемент1.Значение2 <--                             =============}
{============  Элемент2.Значение3                                 =============}
{============  Элемент2.Значение4                                 =============}
{==============================================================================}
function CutLIDE(const PLst: PStringList; const ElName: String): Boolean;
var I: Integer;
begin
    {Инициализация}
    Result:=false;
    If PLst=nil then Exit;

    For I:=PLst^.Count-1 downto 0 do begin
       If CmpStr(CutSlovo(PLst^[I], 1, '.'), ElName)=false then PLst^.Delete(I);
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{===========   Вырезает значения элемента из списка в новый список  ============}
{==============================================================================}
{============  Элемент1.Значение1                                 =============}
{============  Элемент1.Значение2                                 =============}
{============  Элемент2.Значение3 --> Значение3                   =============}
{============  Элемент2.Значение4 --> Значение3                   =============}
{==============================================================================}
function CutLKey(const PSrc, PDest: PStringList; const ElName: String): Boolean;
var I, J  : Integer;
    S, S0 : String;
begin
    {Инициализация}
    Result:=false;
    If (PSrc=nil) or (PDest=nil) then Exit;
    PDest^.Clear;
    S0:=AnsiUpperCase(ElName);

    For I:=0 to PSrc^.Count-1 do begin
       S:=AnsiUpperCase(CutSlovo(PSrc^[I], 1, '.'));
       If S=S0 then begin
          J:=InStrMy(1, PSrc^[I], '.');
          If J>=0 then begin
             S:=PSrc^[I];
             Delete(S, 1, J);
             PDest^.Add(Trim(S));
          end;
       end;
    end;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{================  Вырезает из пути имя файла без расширения  =================}
{==============================================================================}
function ExtractFileNameWithoutExt(const FPath: String): String;
var I: Integer;
begin
    Result:=ExtractFileName(FPath);
    I:=Length(ExtractFileExt(Result));
    Delete(Result, Length(Result)-I+1, I);
end;




{******************************************************************************}
{***  РАБОТА СО СТРОКАМИ С ОГРАНИЧЕНИЯМИ В БЛОКАХ, ОБОЗНАЧЕННЫХ Сh1 и Сh2  ****}
{******************************************************************************}

{==============================================================================}
{======                  ВСПОМОГАТЕЛЬНАЯ ФУНКЦИЯ                      =========}
{==============================================================================}
{======  Замена в строке Str блоков, ограниченных символами Ch1 и Ch2 =========}
{======                  на псевдооператоры: |%N|%                    =========}
{======   Заполняет список соответствий произведенных замен - PsStr   =========}
{==============================================================================}
function ReplPsevdoModul(const Str: String; var PsStr: TStringList;
                         const Ch1, Ch2: Char): String;
var SNum, SBlock: String;
begin
    {Инициализация}
    Result:=Str;

    {Поочередно заменяем все блоки}
    SBlock := CutModulChar(Result, Ch1, Ch2);
    While SBlock<>'' do begin
       SNum   := '|%'+IntToStr(PsStr.Count)+'|%';
       PsStr.Add(SNum+'='+Ch1+SBlock+Ch2);
       Result := ReplModulChar(Result, SNum, Ch1, Ch2);
       SBlock := CutModulChar(Result, Ch1, Ch2);
    end;
end;


{==============================================================================}
{=======  Количество слов в псевдостроке Str, Str0 - разделитель слов  ========}
{==============================================================================}
function GetColPSlovStr(const Str: String; const Str0: String): Integer;
var PsStr : TStringList;
    S     : String;
begin
    {Инициализация}
    Result:=0;
    If (Pos(CH_KAV, Str0)>0) or (Pos('{', Str0)>0) or (Pos('}', Str0)>0)then Exit;
    S     := Str;
    PsStr := TStringList.Create;

    try
       {Заменяем в строке Result блоки псевдо-операторами: |%N|%}
       S:=ReplPsevdoModul(S, PsStr, CH_KAV, CH_KAV);
       S:=ReplPsevdoModul(S, PsStr, '{', '}');

       {Считаем число слов}
       Result:=GetColSlov(S, Str0);

    finally
       PsStr.Free;
    end;
end;


{==============================================================================}
{======  Вырезать с псевдостроки Str слово I, Str0 - разделитель слов  ========}
{==============================================================================}
function CutPSlovoStr(const Str: String; const I: Integer; const Str0: String): String;
var PsStr : TStringList;
    I0    : Integer;
begin
    {Инициализация}
    Result:=Str;
    If (Pos(CH_KAV, Str0)>0) or (Pos('{', Str0)>0) or (Pos('}', Str0)>0)then Exit;
    PsStr := TStringList.Create;

    try
       {Заменяем в строке Result блоки псевдо-операторами: |%N|%}
       Result:=ReplPsevdoModul(Result, PsStr, CH_KAV, CH_KAV);
       Result:=ReplPsevdoModul(Result, PsStr, '{', '}');

       {Вырезаем слово}
       Result:=CutSlovo(Result, I, Str0);

       {Производим в слове обратную замену}
       For I0:=PsStr.Count-1 downto 0 do begin
          Result:=ReplStr(Result, PsStr.Names[I0], PsStr.Values[PsStr.Names[I0]]);
       end;

    finally
       PsStr.Free;
    end;
end;


{==============================================================================}
{=======================  Является ли строка кодом  ===========================}
{==============================================================================}
function IsKod(const SKod: String): Boolean;
var STab, SRow, SCol : String;
    I                : Integer;
begin
    {Инициализация}
    Result := false;

    {Разбиваем код на составляющие}
    If Not SeparatKod(SKod, STab, SRow, SCol) then Exit;
    If (STab='') or (SRow='') or (SCol='') then Exit;

    {Допустимость символов}
    For I:=1 to Length(SKod) do
       If (Not IsIntegerStr(SKod[I])) and (SKod[I]<>CH_SPR) then Exit;

    {Возвращаемый результат}
    Result := true;
end;


{==============================================================================}
{====================  Разбивает код на составляющие  =========================}
{==============================================================================}
function SeparatKod(const SKod: String; var STab, SRow, SCol: String): Boolean;
var S: String;
begin
    {Инициализация}
    Result := false;
    S      := SKod;
    STab   := '';
    SCol   := '';
    SRow   := '';

    {Разбиваем код}
    STab   := CutBlock(S, CH_SPR); If STab <> '' then If Not IsIntegerStr(STab) then Exit;
    SCol   := CutBlock(S, CH_SPR); If SCol <> '' then If Not IsIntegerStr(SCol) then Exit;
    SRow   := CutBlock(S, CH_SPR); If SRow <> '' then If Not IsIntegerStr(SRow) then Exit;

    {Возвращаемый результат}
    Result := true;
end;


{==============================================================================}
{======================  Формирует код из составляющих  =======================}
{==============================================================================}
function GeneratKod(const STab, SRow, SCol: String): String;
begin
    Result := STab+CH_SPR+SCol+CH_SPR+SRow;
end;

end.

