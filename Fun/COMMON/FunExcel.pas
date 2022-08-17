unit FunExcel;

interface

uses  Winapi.Windows,
      System.Variants, System.SysUtils,
      System.Win.ComObj,
      Vcl.Graphics, Vcl.Olectnrs, Vcl.OleServer, Vcl.Clipbrd,
      Excel2000;


{������� ������ EXCEL}
Function CreateExcel:boolean;                                                   StdCall
{���������� ��������� EXCEL}
Function VisibleExcel(visible:boolean):boolean;                                 StdCall
{������� ����� ����� xls}
Function AddXls:boolean;                                                        StdCall
{���������� ��������� �� EXCEL}
Function GetExcelApplicationID:variant;                                         StdCall
{��������� �����}
Function SaveXls: Boolean;                                                      StdCall
{��������� ����� � ���� file_}
Function SaveXlsAs(file_: String): Boolean;                                     StdCall
{������� �����}
Function CloseXls:boolean;                                                      StdCall
{������� Excel}
Function CloseExcel:boolean;                                                    StdCall
{������� ����� SPath}
Function OpenXls(const SPath: String; const IsBlockMacros: Boolean): Boolean;   StdCall
{����������� ������ � ������ �����}
Function StartOfXls:boolean;                                                    StdCall
{����� � ����� ������, ���������� text_ � � ���������}
Function FindTextXls(StrFind: String): String;                                  StdCall
{������� � ������ � �������� findtext_ �� pastetext_}
Function FindAndPasteTextXls(findtext_,pastetext_:string;len_:integer):boolean; StdCall
{������ �������� ������}
Function ReadWBCells(const WB: Variant; const SCell: String): String;           StdCall
Function ReadCells(Pag,X,Y: Integer): String;                                   StdCall
{������������� �������� ������}
Function SetActiveCells(Pag,X,Y: Integer): boolean;                             StdCall
{������ �������� ������ �� �������� ������}
Function MoveFocusActiveCellsRight: Boolean;                                    StdCall
{������ �������� �������� ������}
Function ReadActiveCells: String;                                               StdCall
{����� text_ � �������� ������}
Function WriteActiveCells(text_:string):boolean;                                StdCall

{������������� ������� ���������� �����}
Function SetHeaderXls(Pos: Byte; S:string):boolean;                             StdCall
{������������� ������ ���������� �����}
Function SetFooterXls(Pos: Byte; S:string):boolean;                             StdCall
{������������� ������� �������}
Function SetBackGroundPicXls(FName:string):boolean;                             StdCall

{���� � ����� � ���������� ������-���� � �������� �������}
Function FindKeyTextXls:String;                                                 StdCall

{����� �������� � � �����}
Function GetCountExcelChar(const Ch: Char): Integer;                            StdCall

{������� ������ ������� � �����}
function ExcelColumnNumToChar(const Num: Integer): String;                      StdCall
{������� ���� ������� � �����}
function ExcelColumnCharToNum(const Letters: String): Integer;                  StdCall
{��������� ����� ������ ����� A1 �� ������������}
function ExcelSeparatAddress(const AddressA1: String; var Col, Row: Integer): Boolean; StdCall
{�������� �� ������������ ����� ������ ����� A1}
function ExcelCombineAddress(const Col, Row: Integer): String;                  StdCall
{������������ ����� ������ ����� A1}
function ExcelCorrectAddress(const AddressA1: String; const DCol, DRow: Integer): String; StdCall

var E: Variant;


implementation

uses FunText, FunSys;


{******************************************************************************}
{**************************  ������� ������ EXCEL  ****************************}
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
{***********************  ���������� ��������� EXCEL  *************************}
{******************************************************************************}
function VisibleExcel(Visible: Boolean): Boolean;
begin
    VisibleExcel:=true;
    try    E.Visible:= Visible;
    except VisibleExcel:=false;
    end;
end;


{******************************************************************************}
{**********************  ������� ����� ����� EXCEL  ***************************}
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
{********************  ���������� ��������� �� WORD  **************************}
{******************************************************************************}
function GetExcelApplicationID:variant;
begin
    try    GetExcelApplicationID:=E;
    except GetExcelApplicationID:=0;
    end;
end;


{******************************************************************************}
{*******************************  ��������� �����  ****************************}
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
{************************  ��������� ����� � ���� file_  **********************}
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
{******************************  ������� �����  *******************************}
{******************************************************************************}
function CloseXls: Boolean;
begin
   CloseXls:=true;
   try    E.ActiveWorkbook.Close;
   except CloseXls:=false;
   end;
end;


{******************************************************************************}
{*******************************  ������� Excel  ******************************}
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
{***************************  ������� ����� SPath  ****************************}
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
{*******************  ����������� ������ � ������ �����  **********************}
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
{*********  ����� � ����� ������, ���������� text_ � � ���������  ************}
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
{**********  ������� � ������ � �������� findtext_ �� pastetext_  *************}
{**********  len - ������ ��� ��������� pastetext_                *************}
{******************************************************************************}
function FindAndPasteTextXls(findtext_,pastetext_:string;len_:integer):boolean;
var I, J : Integer;
    S    : String;
    Cel  : array of Char;
begin
    FindAndPasteTextXls:=false;
    try
       {������� ������ �������������� ������}
       S:=FindTextXls(findtext_);
       If S='' then Exit;
       I:=InStrMy(1,S,findtext_);
       If I=0 then Exit;
       {����� ���� ��� ������ ���� � ��������� �����}
       If (len_>1) and (IsIntegerStr(pastetext_)=true) then begin
          SetLength(Cel,len_);
          {��������� ���� ������}
          For I:=Low(Cel) to High(Cel) do Cel[I]:='0';
          {��������� ���� �������}
          J:=Length(pastetext_);
          For I:=High(Cel) downto High(Cel)-Length(pastetext_)+1 do begin
              Cel[I]:=pastetext_[J];
              Dec(J);
          end;
          {���������� ������}
          For I:=Low(Cel) to High(Cel) do begin
             WriteActiveCells(Cel[I]);
             MoveFocusActiveCellsRight;
          end;
       end else begin
          {������� ������}
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
{************************   ������ �������� ������   **************************}
{******************************************************************************}
{************************   SCell = 1!A$1            **************************}
{******************************************************************************}
function ReadWBCells(const WB: Variant; const SCell: String): String;
var SPag, S : String;
    IPag    : Integer;
begin
    {�������������}
    Result := '';
    S      := SCell;
    SPag   := TokChar(S, '!');
    If (Not IsIntegerStr(SPag)) or (S = '') then Exit; IPag := StrToInt(SPag);
    {������ ��������}
    try    If WB.WorkSheets.Count >= IPag then Result:=WB.WorkSheets[IPag].Range[S].Value;
    except end;
end;


{******************************************************************************}
{************************   ������ �������� ������   **************************}
{******************************************************************************}
function ReadCells(Pag,X,Y: Integer): String;
begin
    Result:='';
    If SetActiveCells(Pag,X,Y) = false then Exit;
    Result:=ReadActiveCells;
end;


{******************************************************************************}
{*******************   ������������� �������� ������  *************************}
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
{*******************  ������ �������� ������ �� �������� ������  **************}
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
{*******************  ������ �������� �������� ������  ************************}
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
{*******************  ����� text_ � �������� ������  **************************}
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
{*******************  ������������� ������� ���������� �����  *****************}
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
{*******************  ������������� ������ ���������� �����  ******************}
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
{******************  ������������� ������� �������   **************************}
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
{*********  ���� � ����� � ���������� ������-���� � �������� �������  *********}
{******************************************************************************}
function FindKeyTextXls:String;
const CH1 = '{';
      CH2 = '}';
var Interat: Integer;  {����� �������� ���������}
    MyStart, MyEnd: Integer;
    I : Integer;
    S : String;
begin
    FindKeyTextXls:='';
    try
       S:=FindTextXls(CH1);
       {���������� ������� �������: MyStart - MyEnd}
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
{************************  ����� �������� � � �����  **************************}
{******************************************************************************}
function GetCountExcelChar(const Ch: Char): Integer;
var Pag, PagCount : Integer;
    S, S0         : String;
begin
    {�������������}
    S:='';
    PagCount:=E.ActiveWorkbook.Worksheets.Count; {����� �������� �� ����� +1?}

    {������������� ��� �������� �����}
    For Pag:=PagCount-1 downto 1 do begin
       E.ActiveWorkbook.Worksheets[Pag].Activate;
       E.Cells.Select;         // E.ActiveWorkbook.Worksheets[Pag].Cells.Select;
       E.Cells.Copy;           // E.ActiveWorkbook.Worksheets[Pag].Cells.Copy;
       {������ ��������}
       S0:=Clipboard.AsText;
       S:=S+S0;
       {������� ���������� � �����������}
       E.Application.CutCopyMode:=False;
       E.Range['A1'].Select;
    end;

    {������� ����� �������� � �����}
    Result:=GetColChar(S, Ch);
end;


{******************************************************************************}
{**************      ������� ������ ������� � �����      **********************}
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
{**************       ������� ���� ������� � �����       **********************}
{******************************************************************************}
{**************           'A'=1, 'B'=2, 'AB'=28          **********************}
{******************************************************************************}
function ExcelColumnCharToNum(const Letters: String): Integer;
var S : String;
    I : Byte;
begin
    {�������������}
    Result := 0;
    S      := AnsiUpperCase(Letters);

    {�������� ����������}
    If Not (Length(S) IN [1..2]) then Exit;
    // For I:=1 to Length(S) do If Not (S[I] IN ['A'..'Z']) then Exit;
    For I:=1 to Length(S) do If Not CharInSet(S[I], ['A'..'Z']) then Exit;

    {������}
                        Result :=             Ord(S[1]) - Ord('A') + 1;
    If Length(S)=2 then Result := Result*26 + Ord(S[2]) - Ord('A') + 1;
end;


{******************************************************************************}
{*************   ��������� ����� ������ ����� A1 �� ������������   ************}
{******************************************************************************}
function ExcelSeparatAddress(const AddressA1: String; var Col, Row: Integer): Boolean;
var S1, S2: String;
begin
    {�������������}
    Result := false;
    Col    := 0;
    Row    := 0;
    S2     := AddressA1;

    {������������}
    If Length(S2)<2 then Exit;
    If IsIntegerStr(S2[1]) then Exit;

    {��������� ������ ������}
    S1:=S2[1];
    Delete(S2, 1, 1);

    {��������� ������ ������}
    If Not IsIntegerStr(S2[1]) then begin
       S1:=S1+S2[1];
       Delete(S2, 1, 1);
    end;

    {������������}
    If Length(S2)=0 then Exit;
    If Not IsIntegerStr(S2) then Exit;

    {��������� ���������}
    Row := StrToInt(S2);
    Col := ExcelColumnCharToNum(S1);
    If (Row=0) or (Col=0) then Exit;
    Result:=true;
end;


{******************************************************************************}
{*************   �������� �� ������������ ����� ������ ����� A1    ************}
{******************************************************************************}
function ExcelCombineAddress(const Col, Row: Integer): String;
begin
    {�������������}
    Result:='';

    {������������}
    If (Col<1) or (Row<1) then Exit;

    {��������� ���������}
    Result:=ExcelColumnNumToChar(Col)+IntToStr(Row);
end;


{******************************************************************************}
{*******************   ������������ ����� ������ ����� A1   *******************}
{******************************************************************************}
function ExcelCorrectAddress(const AddressA1: String; const DCol, DRow: Integer): String;
var ICol, IRow: Integer;
begin
    {�������������}
    Result:='';

    {��������� ���������}
    If Not ExcelSeparatAddress(AddressA1, ICol, IRow) then Exit;
    Result:=ExcelCombineAddress(ICol+DCol, IRow+DRow);
end;

end.
