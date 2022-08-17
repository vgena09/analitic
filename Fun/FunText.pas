unit FunText;

interface

uses System.SysUtils, System.Classes, System.Math,
     Vcl.Dialogs, IdGlobal, RxParsing,
     FunType;

{���������� �������� Ch � ������ S}
function  GetColChar(const S: String; const Ch: Char): Integer;
{���������� ����� S0 � ������ S}
function  GetColStr(const S, S0: String): Integer;
{���������� ���� � ������ Str, STerm - ����������� ����}
function  GetColSlov(const Str: String; const STerm: String): Integer;
{�������� �� ������ Str ����� ISlovo, STerm - ����������� ����}
function  CutSlovo(const Str: String; const ISlovo: Integer; const STerm: String): String;
{�������� �� ������ S ����� I � �����, Ch - ����������� ����}
function  CutSlovoEndChar(const S: String; const I: Integer; const Ch: Char): String;
{�������� � ������ S ����� I � ����� �� S0, Ch - ����������� ����}
function  ReplSlovoEndChar(const S, S0: String; const I: Integer; const Ch: Char): String;
{�������� �� ������ Str ILen �������� ������� � ������� IPos}
function  CutStrPos(const Str: String; const IPos, ILen: Integer): String;
{�������� � ������ Str ����� ILen ��������, ������� �� �� ...}
function  CutLongStr(const Str: String; const ILen: Integer): String;
{�������� � ������ Dest ILen �������� ������� � ������� IPos �� ������ SubStr}
function  ReplStrPos(const Dest: String; const IPos, ILen: Integer; const SubStr: String): String;
{�������� � ������ S ��� ��������� S1 �� S2}
function  ReplStr(const S: String; const S1, S2: String): String;
{�������� � ������ S ��� ��������� S1 �� S2 (����������� ����������)}
function  ReplStrLoop(const S: String; const S1, S2: String): String;
{��������� �������� �� ������ S ������ SFind - ����� ������������ �� ��� �������, � ������� Pos}
function  InStrMy(const StartPos:Integer; const S, SFind: String): Integer;
{�������� ���� �� ������ SBlock, ������������ STerm � �������� �������}
function  CutBlock(var SBlock: String; const STerm: String): String;
{�������� ���� �� ������, ������������ ��������� �}
function  TokChar(var Si: String; const C: Char): String;
{�������� ���� �� ������, ������������ ������� S}
function  TokStr(var Si: String; const S: String): String;
{�������� ���� � ����� ������, ������������ ��������� �}
function  TokCharEnd(var Si: String; const C: Char): String;

{�������� �� ������ S1.S2. ... ������ S1}
function  CutClass(var Si: String): String;

{���������� ������� ������� ������� ������� ���������� �����, ������������� ��������� �1-�2}
function  GetPosStartChar(const S: String; const C1, C2: Char): Integer;
{���������� ������� ������� ������� ������� ���������� �����, ������������� �������� S1-S2}
function  GetPosStartStr(const S: String; const S1, S2: String): Integer;
{���������� ������� ������� ������� ������� ������ SPackType}
function  GetPosStartPack(const S, SPackType: String): Integer;
{���������� ������� ���������� ������� ������� ���������� �����, ������������� ��������� �1-�2}
function  GetPosEndChar(const S: String; const StartPos: Integer; const C2:Char): Integer;
{���������� ������� ���������� ������� ������� ���������� �����, ������������� �������� S1-S2}
function  GetPosEndStr(const S: String; const StartPos: Integer; const S2: String): Integer;
{���������� ������� ���������� ������� ������� ������ SPackType}
function  GetPosEndPack(const S, SPackType: String): Integer;

{���������� ��� ������� � ����}
function  SumArrayStr(const S1, S2: array of String): TArrayStr;
{���������� ��� StringList � ����, ������� ������ ���������� ������}
procedure SumSListUniq(const PList, PListAdd: PStringList);
{��������� ������������ � ������}
function  SeparatMStr(const MStr: String; const PLStr: PStringList; const Separator: String): Boolean;
{�������� ������ � ������������}
function  GroupMStr(const PLStr: PStringList; const Separator: String): String;

{�������� ������ ��������� ����, ������������ ��������� �1 � �2}
function  CutModulChar(const S: String; const C1, C2: Char): String;
{�������� ������ ��������� ����, ������������ �������� S1 � S2}
function  CutModulStr(const S: String; const S1, S2: String): String;

{�������� ������ ����� SPack �� ������}
function  CutModulPackList(const Str           : String;
                           const SPackTypeList : array of String;
                           var   SPack         : String): Integer;
{�������� ������ ����� ���� SPackType}
function  CutModulPack(const Str, SPackType: String): String;
{�������� ������ ��������� ����, ������������ ��������� �1 � �2}
function  ReplModulChar(const Str1, Str2: String; const C1, C2: Char): String;
{�������� ������ ��������� ����, ������������ �������� S1 � S2}
function  ReplModulStr(const Str1, Str2: String; const S1, S2: String): String;
{�������� ������ ����� ���� SPackType}
function  ReplModulPack(const Str1, Str2, SPackType: String): String;
{���������� ������ � �������� ����� ��� ������ ���������}
function  CmpStrList(const S: String; SList: array of String): Integer;
{���������� 2 ������ ��� ������ ���������}
function  CmpStr(const S1, S2: String): Boolean;
{����� ������ ���������� �������� ��� ������ ���������}
function  CmpStrFirst(const S1, S2: String): Integer;
{���� ��������� SubStr � ������ Str ��� ����� ��������}
function  FindStr(const SubStr, Str: String): Integer;
{�������� �� ������� �������� �������}
function  IsFormulaDiv(const SFormula: String): Boolean;


{�������� ILine (0, 1, ...) ������ ������������}
function  CutLine(const S: String; const ILine: Integer): String;
{�������� � ����� ������ ������� �hr(10) � Chr(13)}
function  CutEndStr(const S: String): String;
{����������� � ������ Str: ������� � ������� IPos ��������� ����� � ���-�� ICount}
function  UpperStrPart(const Str: String; const IPos, ICount: Integer): String;
{����������� ������ ����� � ���������}
function  UpperArrayStr(const AStr: TArrayStr): TArrayStr;
{����� �������� ������ � ������� �����}
function  FindStrInArray(const StrFind: String; const StrArray: array of String): Boolean;
{����� ��������� ����� � ������� �����}
function  FindIntInArray(const IntFind: Integer; const IntArray: array of Integer): Boolean;

{�������� �� ������ ����� ������}
function  IsIntegerStr(const S: String): Boolean;
{�������� �� ������ ������� ������}
function  IsFloatStr(const S: String): Boolean;
{���������� �������� � ���������� ������� �� ����}
function  CalcStr(const SFormula: String): Extended;

{�������� �������� ������ Section �� ������ PSrc^ � ������ PDest^}
function  CutLSection(const PSrc, PDest: PStringList; const Section: String): Boolean;
{������� �� ������ PLst^ ������ Section}
function  DelLSection(const PLst: PStringList; const Section: String): Boolean;
{��������� � ������ PLst^ ������ �������� ElName.*}
function  CutLIDE(const PLst: PStringList; const ElName: String): Boolean;
{�������� �������� �������� �� ������ � ����� ������}
function  CutLKey(const PSrc, PDest: PStringList; const ElName: String): Boolean;


{�������� �� ���� ��� ����� ��� ����������}
function  ExtractFileNameWithoutExt(const FPath: String): String;


{������ � ������ Str ������, ������������ ��������� Ch1 � Ch2, �� ���������������: |%N|%}
function  ReplPsevdoModul(const Str: String; var PsStr: TStringList; const Ch1, Ch2: Char): String;
{���������� ���� � ������������ Str, Str0 - ����������� ����}
function  GetColPSlovStr(const Str: String; const Str0: String): Integer;
{�������� � ������������ Str ����� I, Str0 - ����������� ����}
function  CutPSlovoStr(const Str: String; const I: Integer; const Str0: String): String;

{�������� �� ������ �����}
function  IsKod(const SKod: String): Boolean;
{��������� ��� �� ������������}
function  SeparatKod(const SKod: String; var STab, SRow, SCol: String): Boolean;
{��������� ��� �� ������������}
function  GeneratKod(const STab, SRow, SCol: String): String;

implementation

uses FunConst;

{$INCLUDE FunText_Modul}


{==============================================================================}
{========================  � � � � �   � � � � � � �  =========================}
{==============================================================================}


{==============================================================================}
{===================   ���������� �������� Ch � ������ S  =====================}
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
{====================   ���������� ����� S0 � ������ S  =======================}
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
{=========  ���������� ���� � ������ Str, STerm - ����������� ����  ===========}
{==============================================================================}
function GetColSlov(const Str: String; const STerm: String): Integer;
begin
    Result:=0;
    While CutSlovo(Str, Result+1, STerm)<>'' do Result:=Result+1;
end;


{==============================================================================}
{======  �������� �� ������ Str ����� ISlovo, STerm - ����������� ����  =======}
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
{======= �������� �� ������ S ����� I c �����, Ch - ����������� ���� ==========}
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
{====== �������� � ������ S ����� I � ����� �� S0, Ch - ����������� ���� ======}
{==============================================================================}
{======   ��� ��������� ������� �� �������������� ������ � ������ �����  ======}
{==============================================================================}
function ReplSlovoEndChar(const S, S0: String; const I: Integer; const Ch: Char): String;
var I0, ILen, ISlovo : Integer;
begin
    {�������������}
    Result := S;
    ISlovo := 0;
    ILen   := 0;
    If Length(Result)=0 then Exit;

    {���� ������ �����}
    For I0:=Length(Result) downto 1 do begin
       If Result[I0]=Ch then begin
          ISlovo := ISlovo+1;
          {����� �������}
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
{======== �������� �� ������ Str ILen �������� ������� � ������� IPos =========}
{==============================================================================}
function CutStrPos(const Str: String; const IPos, ILen: Integer): String;
var I: Integer;
begin
    Result:='';
    If (IPos+ILen-1)>Length(Str) then Exit;
    For I:=IPos to IPos+ILen-1 do Result:=Result+Str[I];
end;


{==============================================================================}
{========  �������� � ������ Str ����� ILen ��������, ������� �� �� ...  ======}
{==============================================================================}
function CutLongStr(const Str: String; const ILen: Integer): String;
begin
    If Length(Str)>ILen then Result:=CutStrPos(Str, 1, ILen-3)+'...'
                        else Result:=Str;
end;


{==============================================================================}
{=�������� � ������ Dest ILen �������� ������� � ������� IPos �� ������ SubStr=}
{==============================================================================}
function ReplStrPos(const Dest: String; const IPos, ILen: Integer; const SubStr: String): String;
begin
    Result:=Dest;
    Delete(Result, IPos, ILen);
    Insert(SubStr, Result, IPos);
end;


{==============================================================================}
{=============  �������� � ������ S ��� ��������� S1 �� S2  ===================}
{==============================================================================}
function ReplStr(const S: String; const S1, S2: String): String;
const MASKREPL = '%|!A%|U!%|!T%|O!%|!D%|O!%|!C%|';
var I: Integer;

    function ReplStrSimpl(const S: String; const S1, S2: String): String;
    begin
        {�������������}
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
{===== �������� � ������ S ��� ��������� S1 �� S2 (����������� ����������) ====}
{==============================================================================}
function ReplStrLoop(const S: String; const S1, S2: String): String;
var S0: String;
begin
    {�������������}
    Result := '';
    S0     := S;

    While S0<>Result do begin
       Result := S0;
       S0     := ReplStr(Result, S1, S2);
    end;
end;


{==============================================================================}
{=============   ��������� �������� �� ������ S ������ SFind    ===============}
{============= ����� ������������ �� ��� �������, � ������� Pos ===============}
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
{==== �������� ���� �� ������ SBlock, ������������ STerm � �������� ������� ===}
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
{============ �������� ���� �� ������, ������������ ��������� � ===============}
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
{============ �������� ���� �� ������, ������������ ������� S =================}
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
{=========  �������� ���� � ����� ������, ������������ ��������� �  ===========}
{==============================================================================}
function TokCharEnd(var Si: String; const C: Char): String;
var I:  Integer;
begin
    Result:=''; if Si='' then Exit;
    I:=Length(Si);
    while (Si[I]<>C)and(I>=1) do begin
        Result:=Si[I]+Result; I:=I-1;
    end;
    {������� ���������� ����}
    Delete(Si, Length(Si)-Length(Result)+1, Length(Result));
    {������� ������-�����������}
    Delete(Si, Length(Si), 1);
end;


{==============================================================================}
{===================  �������� �� ������ S1.S2. ... ������ S1  ================}
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
{======================  ���������� ��� ������� � ����  =======================}
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
{===== ���������� ��� StringList � ����, ������� ������ ���������� ������ =====}
{==============================================================================}
procedure SumSListUniq(const PList, PListAdd: PStringList);
var I: Integer;
begin
    For I:=0 to PListAdd^.Count-1 do begin
       If PList^.IndexOf(PListAdd^[I])<0 then PList^.Add(PListAdd^[I]);
    end;
end;


{==============================================================================}
{=====================  ��������� ������������ � ������  ======================}
{==============================================================================}
function SeparatMStr(const MStr: String; const PLStr: PStringList; const Separator: String): Boolean;
var S: String;
    I: Integer;
begin
    {�������������}
    Result:=false;
    If PLStr=nil then Exit;
    PLStr^.Clear;

    {������������� ��� ������� � ��������� ������}
    I:=1; S:=CutSlovo(MStr, I, Separator);
    While S<>'' do begin
       Result:=true;
       PLStr^.Add(S);
       Inc(I); S:=CutSlovo(MStr, I, Separator);
    end;
end;


{==============================================================================}
{=====================  �������� ������ � ������������   ======================}
{==============================================================================}
function GroupMStr(const PLStr: PStringList; const Separator: String): String;
var I: Integer;
begin
    {�������������}
    Result:='';
    If PLStr=nil then Exit;

    {������������� ��� ������� � ��������� ������}
    For I:=0 to PLStr^.Count-1 do Result:=Result+Separator+PLStr^[I];

    {������� ��������� ������� �����������}
    If Result<>'' then Delete(Result, 1, Length(Separator));
end;


{==============================================================================}
{========  ���������� ������ � �������� ����� ��� ������ ���������  ===========}
{==============================================================================}
function CmpStrList(const S: String; SList: array of String): Integer;
var I  : Integer;
    S0 : String;
begin
    {�������������}
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
{==============  C��������� 2 ������ ��� ������ ���������  ====================}
{==============================================================================}
function CmpStr(const S1, S2: String): Boolean;
begin
    Result:=(AnsiUpperCase(S1)=AnsiUpperCase(S2));
end;


{==============================================================================}
{=========   ����� ������ ���������� �������� ��� ������ ���������   ==========}
{==============================================================================}
function CmpStrFirst(const S1, S2: String): Integer;
var I: Integer;
begin
    {�������������}
    Result:=0;

    For I:=1 to Min(Length(S1), Length(S2)) do begin
       If CmpStr(S1[I], S2[I]) then Result:=I
                               else Break;
    end;
end;



{==============================================================================}
{=================  �������� �� ������� �������� �������  =====================}
{==============================================================================}
function IsFormulaDiv(const SFormula: String): Boolean;
var S, S0 : String;
begin
    {�������������}
    Result := false;
    S      := SFormula;
    S0     := CutModulChar(S, '[', ']');
    {���� �������}
    If S0 <> '' then begin
       While S0 <> '' do begin
          S   := ReplModulChar(S, '', '[', ']');
          S0  := CutModulChar(S, '[', ']');
       end;
       Result := Pos('/', S) > 0;
    end;
end;


{==============================================================================}
{==========  ���� ��������� SubStr � ������ Str ��� ����� ��������  ===========}
{==============================================================================}
function FindStr(const SubStr, Str: String): Integer;
begin
    Result:=Pos(AnsiUpperCase(SubStr), AnsiUpperCase(Str));
end;



{==============================================================================}
{============  �������� ILine (0, 1, ...) ������ ������������  ================}
{==============================================================================}
function CutLine(const S: String; const ILine: Integer): String;
var SList: TStringList;
begin
    {�������������}
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
{==========  �������� � ����� ������ ������� �hr(10) � Chr(13)  ===============}
{==============================================================================}
function CutEndStr(const S: String): String;
begin
    {�������������}
    Result:=S;

    While true do begin
       If Length(Result)=0 then Break;
       If (Result[Length(Result)]<>Chr(10)) and (Result[Length(Result)]<>Chr(13)) then Break;
       Delete(Result, Length(Result), 1);
    end;
end;


{==============================================================================}
{����������� � ������ Str: ������� � ������� IPos ��������� ����� � ���-�� ICount}
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
{==================  ����������� ������ ����� � ���������  ====================}
{==============================================================================}
function UpperArrayStr(const AStr: TArrayStr): TArrayStr;
var I: Integer;
begin
    SetLength(Result, Length(AStr));
    For I:=Low(AStr) to High(AStr) do Result[I]:=AnsiUpperCase(AStr[I]);
end;


{==============================================================================}
{==================  ����� �������� ������ � ������� �����  ===================}
{==============================================================================}
function FindStrInArray(const StrFind: String; const StrArray: array of String): Boolean;
var StrFind0 : String;
    I        : Integer;
begin
    {�������������}
    Result   := false;
    StrFind0 := AnsiUpperCase(StrFind);

    {������������� � ���������� �������� �������}
    For I:=Low(StrArray) to High(StrArray) do begin
       If AnsiUpperCase(StrArray[I])=StrFind0 then begin
          Result:=true;
          Break;
       end;
    end;
end;


{==============================================================================}
{==================  ����� ��������� ����� � ������� �����  ===================}
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
{==================  �������� �� ������ ����� ������  =========================}
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
{==================  �������� �� ������ ������� ������  =======================}
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
{============  ���������� �������� � ���������� ������� �� ����  ==============}
{==============================================================================}
{============                � ������ ������ - 0                 ==============}
{==============================================================================}
function CalcStr(const SFormula: String): Extended;
var SVal, SBlock: String;
begin
    {�������������}
    Result := 0;
    SVal   := ReplStrLoop(SFormula, ' ', '');

    {������������� ����� � ������� �������}
    SBlock := CutModulChar(SVal, '(', ')');
    While SBlock<>'' do begin
       If Pos('/0', SBlock)>0 then Exit;
       SBlock := FloatToStr(GetFormulaValue(SBlock));
       SVal   := ReplModulChar(SVal, SBlock, '(', ')');
       SBlock := CutModulChar(SVal,          '(', ')');
    end;

    {������������� ����������}
    If Pos('/0', SVal)>0 then Exit;
    try    Result := RoundTo(GetFormulaValue(SVal), -4);       // �������� ���������� - ���� ����� �������
    except MessageDlg('������ ���������� ���������:'+CH_NEW+SVal+MSGQWEST, mtError, [mbOk], 0);
    end;

    //Result:=ReplStrLoop(Result, ',', '.');
end;


{==============================================================================}
{=====  �������� �������� ������ Section �� ������ PSrc^ � ������ PDest^  =====}
{==============================================================================}
{============  [������ 1]             =========================================}
{============  ...           --> ...  =========================================}
{============  [������ 2]             =========================================}
{============  ...                    =========================================}
{==============================================================================}
function CutLSection(const PSrc, PDest: PStringList; const Section: String): Boolean;
var I     : Integer;
    B     : Boolean;
    S, S0 : String;
begin
    {�������������}
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
{============  ������� �� ������ PLst^ ������ Section  ========================}
{==============================================================================}
{============  [������ 1]        X    =========================================}
{============  ...               X    =========================================}
{============  [������ 2]             =========================================}
{============  ...                    =========================================}
{==============================================================================}
function DelLSection(const PLst: PStringList; const Section: String): Boolean;
label Nx;
var I     : Integer;
    B     : Boolean;
    S, S0 : String;
begin
    {�������������}
    Result := false;
    If PLst=nil then Exit;
    S0     := AnsiUpperCase(Section);
    B      := false;
    I      := 0;

Nx: If I<=PLst^.Count-1 then begin
       {��������� �������� ������}
       S:=Trim(CutModulChar(Trim(PLst^[I]), '[', ']'));
       If S<>'' then begin
          If AnsiUpperCase(S)=S0 then B:=true else B:=false;
       end;
       {������� ������ ��� ��������� � ���������}
       If B=true then PLst^.Delete(I)
                 else I:=I+1;
       {��������� ������}
       Goto Nx;
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{===========   ��������� � ������ PLst^ ������ �������� ElName.*  =============}
{==============================================================================}
{============  �������1.��������1 <-- ElName=�������1             =============}
{============  �������1.��������2 <--                             =============}
{============  �������2.��������3                                 =============}
{============  �������2.��������4                                 =============}
{==============================================================================}
function CutLIDE(const PLst: PStringList; const ElName: String): Boolean;
var I: Integer;
begin
    {�������������}
    Result:=false;
    If PLst=nil then Exit;

    For I:=PLst^.Count-1 downto 0 do begin
       If CmpStr(CutSlovo(PLst^[I], 1, '.'), ElName)=false then PLst^.Delete(I);
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{===========   �������� �������� �������� �� ������ � ����� ������  ============}
{==============================================================================}
{============  �������1.��������1                                 =============}
{============  �������1.��������2                                 =============}
{============  �������2.��������3 --> ��������3                   =============}
{============  �������2.��������4 --> ��������3                   =============}
{==============================================================================}
function CutLKey(const PSrc, PDest: PStringList; const ElName: String): Boolean;
var I, J  : Integer;
    S, S0 : String;
begin
    {�������������}
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

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{================  �������� �� ���� ��� ����� ��� ����������  =================}
{==============================================================================}
function ExtractFileNameWithoutExt(const FPath: String): String;
var I: Integer;
begin
    Result:=ExtractFileName(FPath);
    I:=Length(ExtractFileExt(Result));
    Delete(Result, Length(Result)-I+1, I);
end;




{******************************************************************************}
{***  ������ �� �������� � ������������� � ������, ������������ �h1 � �h2  ****}
{******************************************************************************}

{==============================================================================}
{======                  ��������������� �������                      =========}
{==============================================================================}
{======  ������ � ������ Str ������, ������������ ��������� Ch1 � Ch2 =========}
{======                  �� ���������������: |%N|%                    =========}
{======   ��������� ������ ������������ ������������� ����� - PsStr   =========}
{==============================================================================}
function ReplPsevdoModul(const Str: String; var PsStr: TStringList;
                         const Ch1, Ch2: Char): String;
var SNum, SBlock: String;
begin
    {�������������}
    Result:=Str;

    {���������� �������� ��� �����}
    SBlock := CutModulChar(Result, Ch1, Ch2);
    While SBlock<>'' do begin
       SNum   := '|%'+IntToStr(PsStr.Count)+'|%';
       PsStr.Add(SNum+'='+Ch1+SBlock+Ch2);
       Result := ReplModulChar(Result, SNum, Ch1, Ch2);
       SBlock := CutModulChar(Result, Ch1, Ch2);
    end;
end;


{==============================================================================}
{=======  ���������� ���� � ������������ Str, Str0 - ����������� ����  ========}
{==============================================================================}
function GetColPSlovStr(const Str: String; const Str0: String): Integer;
var PsStr : TStringList;
    S     : String;
begin
    {�������������}
    Result:=0;
    If (Pos(CH_KAV, Str0)>0) or (Pos('{', Str0)>0) or (Pos('}', Str0)>0)then Exit;
    S     := Str;
    PsStr := TStringList.Create;

    try
       {�������� � ������ Result ����� ������-�����������: |%N|%}
       S:=ReplPsevdoModul(S, PsStr, CH_KAV, CH_KAV);
       S:=ReplPsevdoModul(S, PsStr, '{', '}');

       {������� ����� ����}
       Result:=GetColSlov(S, Str0);

    finally
       PsStr.Free;
    end;
end;


{==============================================================================}
{======  �������� � ������������ Str ����� I, Str0 - ����������� ����  ========}
{==============================================================================}
function CutPSlovoStr(const Str: String; const I: Integer; const Str0: String): String;
var PsStr : TStringList;
    I0    : Integer;
begin
    {�������������}
    Result:=Str;
    If (Pos(CH_KAV, Str0)>0) or (Pos('{', Str0)>0) or (Pos('}', Str0)>0)then Exit;
    PsStr := TStringList.Create;

    try
       {�������� � ������ Result ����� ������-�����������: |%N|%}
       Result:=ReplPsevdoModul(Result, PsStr, CH_KAV, CH_KAV);
       Result:=ReplPsevdoModul(Result, PsStr, '{', '}');

       {�������� �����}
       Result:=CutSlovo(Result, I, Str0);

       {���������� � ����� �������� ������}
       For I0:=PsStr.Count-1 downto 0 do begin
          Result:=ReplStr(Result, PsStr.Names[I0], PsStr.Values[PsStr.Names[I0]]);
       end;

    finally
       PsStr.Free;
    end;
end;


{==============================================================================}
{=======================  �������� �� ������ �����  ===========================}
{==============================================================================}
function IsKod(const SKod: String): Boolean;
var STab, SRow, SCol : String;
    I                : Integer;
begin
    {�������������}
    Result := false;

    {��������� ��� �� ������������}
    If Not SeparatKod(SKod, STab, SRow, SCol) then Exit;
    If (STab='') or (SRow='') or (SCol='') then Exit;

    {������������ ��������}
    For I:=1 to Length(SKod) do
       If (Not IsIntegerStr(SKod[I])) and (SKod[I]<>CH_SPR) then Exit;

    {������������ ���������}
    Result := true;
end;


{==============================================================================}
{====================  ��������� ��� �� ������������  =========================}
{==============================================================================}
function SeparatKod(const SKod: String; var STab, SRow, SCol: String): Boolean;
var S: String;
begin
    {�������������}
    Result := false;
    S      := SKod;
    STab   := '';
    SCol   := '';
    SRow   := '';

    {��������� ���}
    STab   := CutBlock(S, CH_SPR); If STab <> '' then If Not IsIntegerStr(STab) then Exit;
    SCol   := CutBlock(S, CH_SPR); If SCol <> '' then If Not IsIntegerStr(SCol) then Exit;
    SRow   := CutBlock(S, CH_SPR); If SRow <> '' then If Not IsIntegerStr(SRow) then Exit;

    {������������ ���������}
    Result := true;
end;


{==============================================================================}
{======================  ��������� ��� �� ������������  =======================}
{==============================================================================}
function GeneratKod(const STab, SRow, SCol: String): String;
begin
    Result := STab+CH_SPR+SCol+CH_SPR+SRow;
end;

end.

