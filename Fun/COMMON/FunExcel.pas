unit FunExcel;

interface

uses  Winapi.Windows,
      System.Variants, System.SysUtils,
      System.Win.ComObj,
      Vcl.Graphics, Vcl.Olectnrs, Vcl.OleServer, Vcl.Clipbrd,
      Excel2000;


{Создать объект EXCEL}
Function CreateExcel:boolean;                                                   StdCall
{Установить видимость EXCEL}
Function VisibleExcel(visible:boolean):boolean;                                 StdCall
{Создать новую книгу xls}
Function AddXls:boolean;                                                        StdCall
{Возвратить указатель на EXCEL}
Function GetExcelApplicationID:variant;                                         StdCall
{Сохранить книгу}
Function SaveXls: Boolean;                                                      StdCall
{Сохранить книгу в файл file_}
Function SaveXlsAs(file_: String): Boolean;                                     StdCall
{Закрыть книгу}
Function CloseXls:boolean;                                                      StdCall
{Закрыть Excel}
Function CloseExcel:boolean;                                                    StdCall
{Открыть книгу SPath}
Function OpenXls(const SPath: String; const IsBlockMacros: Boolean): Boolean;   StdCall
{Переместить курсор в начало книги}
Function StartOfXls:boolean;                                                    StdCall
{Поиск в книге ячейки, содержащей text_ и её выделение}
Function FindTextXls(StrFind: String): String;                                  StdCall
{Находит в тексте и заменяет findtext_ на pastetext_}
Function FindAndPasteTextXls(findtext_,pastetext_:string;len_:integer):boolean; StdCall
{Читает значение ячейки}
Function ReadWBCells(const WB: Variant; const SCell: String): String;           StdCall
Function ReadCells(Pag,X,Y: Integer): String;                                   StdCall
{Устанавливает активную ячейку}
Function SetActiveCells(Pag,X,Y: Integer): boolean;                             StdCall
{Меняет активную ячейку на соседнюю справа}
Function MoveFocusActiveCellsRight: Boolean;                                    StdCall
{Читает значение активной ячейки}
Function ReadActiveCells: String;                                               StdCall
{Пишет text_ в активную ячейку}
Function WriteActiveCells(text_:string):boolean;                                StdCall

{Устанавливает верхний колонтитул книги}
Function SetHeaderXls(Pos: Byte; S:string):boolean;                             StdCall
{Устанавливает нижний колонтитул книги}
Function SetFooterXls(Pos: Byte; S:string):boolean;                             StdCall
{Устанавливает фоновый рисунок}
Function SetBackGroundPicXls(FName:string):boolean;                             StdCall

{Ищет в книге и возвращает строку-ключ в фигурных скобках}
Function FindKeyTextXls:String;                                                 StdCall

{Число символов С в книге}
Function GetCountExcelChar(const Ch: Char): Integer;                            StdCall

{Перевод номера столбца в буквы}
function ExcelColumnNumToChar(const Num: Integer): String;                      StdCall
{Перевод букв столбца в номер}
function ExcelColumnCharToNum(const Letters: String): Integer;                  StdCall
{Разделяет адрес ячейки формы A1 на составляющие}
function ExcelSeparatAddress(const AddressA1: String; var Col, Row: Integer): Boolean; StdCall
{Собирает из составляющих адрес ячейки формы A1}
function ExcelCombineAddress(const Col, Row: Integer): String;                  StdCall
{Корректирует адрес ячейки формы A1}
function ExcelCorrectAddress(const AddressA1: String; const DCol, DRow: Integer): String; StdCall

var E: Variant;


implementation

uses FunText, FunSys;


{******************************************************************************}
{**************************  Создать объект EXCEL  ****************************}
{******************************************************************************}
function CreateExcel:boolean;
begin
    CreateExcel:=true;
    try    E:=CreateOleObject('Excel.Application');
    except CreateExcel:=false;
           E:=Unassigned;
    end;
end;


{******************************************************************************}
{***********************  Установить видимость EXCEL  *************************}
{******************************************************************************}
function VisibleExcel(Visible: Boolean): Boolean;
begin
    VisibleExcel:=true;
    try    E.Visible:= Visible;
    except VisibleExcel:=false;
    end;
end;


{******************************************************************************}
{**********************  Создать новую книгу EXCEL  ***************************}
{******************************************************************************}
function AddXls:boolean;
var Xls_: Variant;
begin
    AddXls:=true;
    try    Xls_:=E.Workbooks; //.Documents;
           Xls_.Add;
    except AddXls:=false;
    end;
end;


{******************************************************************************}
{********************  Возвратить указатель на WORD  **************************}
{******************************************************************************}
function GetExcelApplicationID:variant;
begin
    try    GetExcelApplicationID:=E;
    except GetExcelApplicationID:=0;
    end;
end;


{******************************************************************************}
{*******************************  Сохранить книгу  ****************************}
{******************************************************************************}
function SaveXls: Boolean;
var B: Boolean;
begin
    SaveXls := true;
    try
       B:=E.Application.DisplayAlerts;
       E.Application.DisplayAlerts:=false;
       E.ActiveWorkbook.Save;
    except
       SaveXls:=false;
       Exit;
    end;
    E.Application.DisplayAlerts:=B;
end;


{******************************************************************************}
{************************  Сохранить книгу в файл file_  **********************}
{******************************************************************************}
function SaveXlsAs(file_: String): Boolean;
var B: Boolean;
begin
    SaveXlsAs:=true;
    try
       ForceDirectories(ExtractFilePath(file_));
       B:=E.Application.DisplayAlerts;
       E.Application.DisplayAlerts:=false;
       E.ActiveWorkbook.SaveAs(file_);
    except
       SaveXlsAs:=false;
       Exit;
    end;
    E.Application.DisplayAlerts:=B;
end;


{******************************************************************************}
{******************************  Закрыть книгу  *******************************}
{******************************************************************************}
function CloseXls: Boolean;
begin
   CloseXls:=true;
   try    E.ActiveWorkbook.Close;
   except CloseXls:=false;
   end;
end;


{******************************************************************************}
{*******************************  Закрыть Excel  ******************************}
{******************************************************************************}
function CloseExcel: Boolean;
begin
   CloseExcel:=true;
   try    E.DisplayAlerts:=false;
          E.Application.Quit;
          E.Quit;
          E:=Unassigned;
   except CloseExcel:=false;
   end;
end;


{******************************************************************************}
{***************************  Открыть книгу SPath  ****************************}
{******************************************************************************}
function OpenXls(const SPath: String; const IsBlockMacros: Boolean): Boolean;
begin
   Result:=true;
   try
      If IsBlockMacros then begin
         E.DisplayAlerts := false;
         E.EnableEvents  := false;
      end;
      E.Workbooks.Open(SPath);
   except
      Result:=false;
   end;
end;


{******************************************************************************}
{*******************  Переместить курсор в начало книги  **********************}
{******************************************************************************}
function StartOfXls:boolean;
begin
    StartOfXls:=true;
    try
       SetActiveCells(1,1,1);
    except
       StartOfXls:=false;
    end;
end;


{******************************************************************************}
{*********  Поиск в книге ячейки, содержащей text_ и её выделение  ************}
{******************************************************************************}
function FindTextXls(StrFind: String): String;
var Pag,PagCount : Integer;
    S : String;
    R : ExcelRange;
begin
    FindTextXls:='';
    try
       PagCount:=E.ActiveWorkbook.Worksheets.Count;
       For Pag:=1 to PagCount do begin
          E.ActiveWorkbook.Worksheets[Pag].Activate;
          E.ActiveSheet.Cells.Item[1,1].Activate;
          IDispatch(R):=E.ActiveCell.Find(What := StrFind, LookIn := xlValues, SearchDirection := xlNext);
          If not Assigned(R) then Continue
                             else R.Activate;
          S:=ReadActiveCells;
          If InStrMy(1, S, StrFind)>0 then begin
             FindTextXls:=S;
             Exit;
          end;
       end;
    except
       FindTextXls:='';
    end;
end;


{******************************************************************************}
{**********  Находит в тексте и заменяет findtext_ на pastetext_  *************}
{**********  len - только для числового pastetext_                *************}
{******************************************************************************}
function FindAndPasteTextXls(findtext_,pastetext_:string;len_:integer):boolean;
var I, J : Integer;
    S    : String;
    Cel  : array of Char;
begin
    FindAndPasteTextXls:=false;
    try
       {Находим первую корректируемую ячейку}
       S:=FindTextXls(findtext_);
       If S='' then Exit;
       I:=InStrMy(1,S,findtext_);
       If I=0 then Exit;
       {Может быть это запись кода в несколько ячеек}
       If (len_>1) and (IsIntegerStr(pastetext_)=true) then begin
          SetLength(Cel,len_);
          {Заполняем поле нолями}
          For I:=Low(Cel) to High(Cel) do Cel[I]:='0';
          {Заполняем поле цифрами}
          J:=Length(pastetext_);
          For I:=High(Cel) downto High(Cel)-Length(pastetext_)+1 do begin
              Cel[I]:=pastetext_[J];
              Dec(J);
          end;
          {Производим запись}
          For I:=Low(Cel) to High(Cel) do begin
             WriteActiveCells(Cel[I]);
             MoveFocusActiveCellsRight;
          end;
       end else begin
          {Обычная замена}
          Delete(S,I,Length(findtext_));
          Insert(pastetext_,S,I);
          WriteActiveCells(S);
       end;
       FindAndPasteTextXls:=true;
    except
       FindAndPasteTextXls:=false;
    end;
end;


{******************************************************************************}
{************************   Читает значение ячейки   **************************}
{******************************************************************************}
{************************   SCell = 1!A$1            **************************}
{******************************************************************************}
function ReadWBCells(const WB: Variant; const SCell: String): String;
var SPag, S : String;
    IPag    : Integer;
begin
    {Инициализация}
    Result := '';
    S      := SCell;
    SPag   := TokChar(S, '!');
    If (Not IsIntegerStr(SPag)) or (S = '') then Exit; IPag := StrToInt(SPag);
    {Читаем значение}
    try    If WB.WorkSheets.Count >= IPag then Result:=WB.WorkSheets[IPag].Range[S].Value;
    except end;
end;


{******************************************************************************}
{************************   Читает значение ячейки   **************************}
{******************************************************************************}
function ReadCells(Pag,X,Y: Integer): String;
begin
    Result:='';
    If SetActiveCells(Pag,X,Y) = false then Exit;
    Result:=ReadActiveCells;
end;


{******************************************************************************}
{*******************   Устанавливает активную ячейку  *************************}
{******************************************************************************}
function SetActiveCells(Pag,X,Y: Integer): boolean;
begin
    try
       SetActiveCells:=true;
       E.ActiveWorkbook.Worksheets[Pag].Activate;
       E.ActiveSheet.Cells.Item[X,Y].Activate;
    except
       SetActiveCells:=false;
    end;
end;


{******************************************************************************}
{*******************  Меняет активную ячейку на соседнюю справа  **************}
{******************************************************************************}
function MoveFocusActiveCellsRight: Boolean;
begin
    try
       E.ActiveCell.Offset[0,1].Activate;    //.Item[-1,0]
       MoveFocusActiveCellsRight:=true;
    except
       MoveFocusActiveCellsRight:=false;
    end;
end;


{******************************************************************************}
{*******************  Читает значение активной ячейки  ************************}
{******************************************************************************}
function ReadActiveCells: String;
begin
    try
       ReadActiveCells:=E.ActiveCell.Value;
    except
       ReadActiveCells:='';
    end;
end;


{******************************************************************************}
{*******************  Пишет text_ в активную ячейку  **************************}
{******************************************************************************}
function WriteActiveCells(text_:string):boolean;
begin
    try
       WriteActiveCells:=true;
       E.ActiveCell.Value:=text_;
    except
       WriteActiveCells:=false;
    end;
end;


{******************************************************************************}
{*******************  Устанавливает верхний колонтитул книги  *****************}
{******************************************************************************}
function SetHeaderXls(Pos: Byte; S:string):boolean;
var Pag, PagCount: Integer;
begin
    SetHeaderXls:=true;
    try
    PagCount:=E.ActiveWorkbook.Worksheets.Count;
    For Pag:=1 to PagCount do begin
       Case Pos of
       1: E.ActiveWorkbook.Worksheets[Pag].PageSetup.LeftHeader:=S;
       2: E.ActiveWorkbook.Worksheets[Pag].PageSetup.CenterHeader:=S;
       3: E.ActiveWorkbook.Worksheets[Pag].PageSetup.RightHeader:=S;
       end;
    end;
    except
       SetHeaderXls:=false;
    end;
end;


{******************************************************************************}
{*******************  Устанавливает нижний колонтитул книги  ******************}
{******************************************************************************}
Function SetFooterXls(Pos: Byte; S:string):boolean;
var Pag,PagCount: Integer;
begin
    SetFooterXls:=true;
    try
    PagCount:=E.ActiveWorkbook.Worksheets.Count;
    For Pag:=1 to PagCount do begin
       Case Pos of
       1: E.ActiveWorkbook.Worksheets[Pag].PageSetup.LeftFooter:=S;
       2: E.ActiveWorkbook.Worksheets[Pag].PageSetup.CenterFooter:=S;
       3: E.ActiveWorkbook.Worksheets[Pag].PageSetup.RightFooter:=S;
       end;
    end;
    except
       SetFooterXls:=false;
    end;
end;


{******************************************************************************}
{******************  Устанавливает фоновый рисунок   **************************}
{******************************************************************************}
Function SetBackGroundPicXls(FName:string):boolean;
var Pag,PagCount: Integer;
begin
    SetBackGroundPicXls:=true;
    try
       PagCount:=E.ActiveWorkbook.Worksheets.Count;
       For Pag:=1 to PagCount do E.ActiveWorkbook.Worksheets[Pag].SetBackgroundPicture(FName);
    except
       SetBackGroundPicXls:=false;
    end;
end;


{******************************************************************************}
{*********  Ищет в книге и возвращает строку-ключ в фигурных скобках  *********}
{******************************************************************************}
function FindKeyTextXls:String;
const CH1 = '{';
      CH2 = '}';
var Interat: Integer;  {Число вложеных процессов}
    MyStart, MyEnd: Integer;
    I : Integer;
    S : String;
begin
    FindKeyTextXls:='';
    try
       S:=FindTextXls(CH1);
       {Определяем границы области: MyStart - MyEnd}
       MyStart:=InStrMy(1,S,CH1);
       If MyStart=0 then Exit;
       Interat:=0;
       For I:=MyStart to Length(S) do begin
          If S[I]=CH1 then Inc(Interat);
          If S[I]=CH2 then Dec(Interat);
          If Interat=0 then break;
       end;
       If Interat<>0 then Exit;
       MyEnd:=I;
       Delete(S,1,MyStart-1);
       Delete(S,MyEnd+1,Length(S));
       FindKeyTextXls:=S;
    except
       FindKeyTextXls:='';
    end;
end;


{******************************************************************************}
{************************  Число символов С в книге  **************************}
{******************************************************************************}
function GetCountExcelChar(const Ch: Char): Integer;
var Pag, PagCount : Integer;
    S, S0         : String;
begin
    {Инициализация}
    S:='';
    PagCount:=E.ActiveWorkbook.Worksheets.Count; {Может работать не верно +1?}

    {Просматриваем все страницы книги}
    For Pag:=PagCount-1 downto 1 do begin
       E.ActiveWorkbook.Worksheets[Pag].Activate;
       E.Cells.Select;         // E.ActiveWorkbook.Worksheets[Pag].Cells.Select;
       E.Cells.Copy;           // E.ActiveWorkbook.Worksheets[Pag].Cells.Copy;
       {Читаем клипборд}
       S0:=Clipboard.AsText;
       S:=S+S0;
       {Снимаем готовность к копированию}
       E.Application.CutCopyMode:=False;
       E.Range['A1'].Select;
    end;

    {Считаем число символов в книге}
    Result:=GetColChar(S, Ch);
end;


{******************************************************************************}
{**************      Перевод номера столбца в буквы      **********************}
{******************************************************************************}
{**************           1='A', 2='B', 28='AB'          **********************}
{******************************************************************************}
function ExcelColumnNumToChar(const Num: Integer): String;
// var I: Integer;
begin
    If Num > 26 then begin
       Result := Chr(((Num - 1) Div 26) + 64)+Chr(((Num - 1) Mod 26) + 65);
    end else begin
       Result := Chr(Num + 64);
    end;
//    I := Trunc(Num / 26);
//    If Num > 26 then Result := Chr(I + 64) + Chr(Num - (I * 26) + 64)
//                else Result := Chr(Num + 64);
end;


{******************************************************************************}
{**************       Перевод букв столбца в номер       **********************}
{******************************************************************************}
{**************           'A'=1, 'B'=2, 'AB'=28          **********************}
{******************************************************************************}
function ExcelColumnCharToNum(const Letters: String): Integer;
var S : String;
    I : Byte;
begin
    {Инициализация}
    Result := 0;
    S      := AnsiUpperCase(Letters);

    {Проверка параметров}
    If Not (Length(S) IN [1..2]) then Exit;
    // For I:=1 to Length(S) do If Not (S[I] IN ['A'..'Z']) then Exit;
    For I:=1 to Length(S) do If Not CharInSet(S[I], ['A'..'Z']) then Exit;

    {Расчет}
                        Result :=             Ord(S[1]) - Ord('A') + 1;
    If Length(S)=2 then Result := Result*26 + Ord(S[2]) - Ord('A') + 1;
end;


{******************************************************************************}
{*************   Разделяет адрес ячейки формы A1 на составляющие   ************}
{******************************************************************************}
function ExcelSeparatAddress(const AddressA1: String; var Col, Row: Integer): Boolean;
var S1, S2: String;
begin
    {Инициализация}
    Result := false;
    Col    := 0;
    Row    := 0;
    S2     := AddressA1;

    {Допустимость}
    If Length(S2)<2 then Exit;
    If IsIntegerStr(S2[1]) then Exit;

    {Переносим первый символ}
    S1:=S2[1];
    Delete(S2, 1, 1);

    {Переносим второй символ}
    If Not IsIntegerStr(S2[1]) then begin
       S1:=S1+S2[1];
       Delete(S2, 1, 1);
    end;

    {Допустимость}
    If Length(S2)=0 then Exit;
    If Not IsIntegerStr(S2) then Exit;

    {Формируем результат}
    Row := StrToInt(S2);
    Col := ExcelColumnCharToNum(S1);
    If (Row=0) or (Col=0) then Exit;
    Result:=true;
end;


{******************************************************************************}
{*************   Собирает из составляющих адрес ячейки формы A1    ************}
{******************************************************************************}
function ExcelCombineAddress(const Col, Row: Integer): String;
begin
    {Инициализация}
    Result:='';

    {Допустимость}
    If (Col<1) or (Row<1) then Exit;

    {Формируем результат}
    Result:=ExcelColumnNumToChar(Col)+IntToStr(Row);
end;


{******************************************************************************}
{*******************   Корректирует адрес ячейки формы A1   *******************}
{******************************************************************************}
function ExcelCorrectAddress(const AddressA1: String; const DCol, DRow: Integer): String;
var ICol, IRow: Integer;
begin
    {Инициализация}
    Result:='';

    {Формируем результат}
    If Not ExcelSeparatAddress(AddressA1, ICol, IRow) then Exit;
    Result:=ExcelCombineAddress(ICol+DCol, IRow+DRow);
end;

end.
