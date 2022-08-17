{******************************************************************************}
{******************      ��������������� �������          *********************}
{******************  �� ������ � ������������ ���������   *********************}
{******************************************************************************}

{$INCLUDE MAIN_VIRTUAL_INI}

{==============================================================================}
{===================   ����������� ��������� �������   ========================}
{==============================================================================}
{===================  IsShort - �� ��������� �������   ========================}
{==============================================================================}
function TFMAIN.FormulaToCaption(const SFormula: String; const IsShort: Boolean): String;
var Str, S1, S2: String;
begin
    {�������������}
    Result := '';
    S1 := Trim(CutModulChar(SFormula, '[', ']'));

    {������� �������}
    If S1='' then begin
       Result := KeyToCaption(SFormula, CH_SPR, ' / ', IsShort);   //CHR(10)

    {������� �������}
    end else begin
       {�������� ��������}
       S1  := KeyToCaption(S1, CH_SPR, ' / ', IsShort);
       Str := ReplModulChar(SFormula, STR_UNIQ, '[', ']');
       S2  := Trim(CutModulChar(Str, '[', ']'));
       If S2<>'' then S2 := KeyToCaption(S2, CH_SPR, ' / ', IsShort);

       {��������� ���������}
       Result := ReplModulChar(Str, '[Y]', '[', ']');
       Result := ReplStr(Result, STR_UNIQ, '[X]')+CHR(10);
       Result := Result+'X = '+S1;
       If S2<>'' then Result := Result+Chr(10)+'Y = '+S2;
    end;
end;

function TFMAIN.KeyToCaption(const SKey, SSeparatKey, SSeparatResult: String; const IsShort: Boolean): String;
var SCol, SRow, SKey0, STab: String;
    ITab : Integer;
begin
    {�������������}
    Result  := '';
    SKey0   := SKey;

    try
       STab := CutBlock(SKey0, SSeparatKey);
       If Not IsIntegerStr(STab) then Exit;
       ITab := StrToInt(STab);
       If Not IsShort then begin
          STab := CutLongStr(TablesCounterToCaption(STab), 100);
          If STab='' then Exit;
          Result := STab;
       end;

       SCol := CutLongStr(ColsNumericToCaption(CutBlock(SKey0, SSeparatKey), ITab), 100);
       If SCol <> '' then begin
          If IsShort then Result := SCol
                     else Result := Result+SSeparatResult + SCol;
       end;

       SRow := CutLongStr(RowsNumericToCaption(CutBlock(SKey0, SSeparatKey), ITab), 100);
       If (SRow <> '') and (SCol <> SRow) then Result := Result + SSeparatResult + SRow;
    finally
    end;
end;

function TFMAIN.RegionsCounterToCaption(const FCounter: Integer): String;
var LRegion_ : TLRegion;
begin
    Result := '';
    try     If Not FindRegions(@LRegion_, true, FCounter, '', -1, -1, 0) then Exit;
            Result := LRegion_[0].FCaption;
    finally SetLength(LRegion_, 0);
    end;
end;

function TFMAIN.RegionsCounterToParent(const FCounter: Integer): Integer;
var LRegion_ : TLRegion;
begin
    Result := -1;
    try     If Not FindRegions(@LRegion_, true, FCounter, '', -1, -1, 0) then Exit;
            Result := LRegion_[0].FParent;
    finally SetLength(LRegion_, 0);
    end;
end;

function TFMAIN.TablesCounterToBankPrefix(const FCounter: String; const Dat: TDate): String;
var LTable_ : TLTable;
begin
    If IMainBank = -1 then begin
       {��������� ����� ������}
       Result := '';
       try     If Not FindTables(@LTable_, true, FCounter, '', Dat) then Exit;
               Result := LTable_[0].SBankPref;
       finally SetLength(LTable_, 0);
       end;
    end else begin
       {�������������� ��������� ����� ������}
       Result := LBank[IMainBank].Pref;
    end;
end;

function TFMAIN.TablesCounterToCaption(const FCounter: String): String;
var LTable_ : TLTable;
begin
    Result := '';
    try     If Not FindTables(@LTable_, true, FCounter, '', 0) then Exit;
            Result := LTable_[0].FCaption;
    finally SetLength(LTable_, 0);
    end;
end;

function TFMAIN.IsTableNoStandart(const FCounter: String): Boolean;
var LTable_ : TLTable;
begin
    Result := false;
    try     If Not FindTables(@LTable_, true, FCounter, '', 0) then Exit;
            Result := LTable_[0].FNoStand;
    finally SetLength(LTable_, 0);
    end;
end;

function TFMAIN.RowsNumericToCaption(const FNumeric: String; const FTable: Integer): String;
var LRow_ : TLRow;
begin
    Result := '';
    try     If Not FindRows(@LRow_, true, FNumeric, '', FTable) then Exit;
            Result := LRow_[0].FCaption;
    finally SetLength(LRow_, 0);
    end;
end;

function TFMAIN.ColsNumericToCaption(const FNumeric: String; const FTable: Integer): String;
var LCol_ : TLCol;
begin
    Result := '';
    try     If Not FindCols(@LCol_, true, FNumeric, '', FTable) then Exit;
            Result := LCol_[0].FCaption;
    finally SetLength(LCol_, 0);
    end;
end;

function TFMAIN.FormsCounterToBankPrefix(const FCounter: Integer): String;
var LForma_ : TLForma;
begin
    Result := '';
    try     If Not FindForms(@LForma_, true, false, FCounter, -1, '', '', '', '', false, 0) then Exit;
            Result := LForma_[0].SBankPref;
    finally SetLength(LForma_, 0);
    end;
end;

function TFMAIN.FormsCounterToCaption(const FCounter: Integer): String;
var LForma_ : TLForma;
begin
    Result := '';
    try     If Not FindForms(@LForma_, true, false, FCounter, -1, '', '', '', '', false, 0) then Exit;
            Result := LForma_[0].FCaption;
    finally SetLength(LForma_, 0);
    end;
end;

function TFMAIN.FormsCounterToFile(const FCounter: Integer): String;
var LForma_ : TLForma;
begin
    Result := '';
    try     If Not FindForms(@LForma_, true, false, FCounter, -1, '', '', '', '', false, 0) then Exit;
            Result := PathFileForm(LForma_[0].SBankPref, LForma_[0].FFile);
    finally SetLength(LForma_, 0);
    end;
end;

function TFMAIN.FormsCounterToParent(const FCounter: Integer): Integer;
var LForma_ : TLForma;
begin
    Result := -1;
    try     If Not FindForms(@LForma_, true, false, FCounter, -1, '', '', '', '', false, 0) then Exit;
            Result := LForma_[0].FParent;
    finally SetLength(LForma_, 0);
    end;
end;

function TFMAIN.FormsCounterToPathTreeCaption(const FCounter: Integer): String;
var IParent, IChild : Integer;
begin
    Result  := '';
    IChild  := FCounter;
    IParent := FormsCounterToParent(IChild);
    While IParent >= 0 do begin
       Result  := FormsCounterToCaption(IChild) + '\' + Result;
       IChild  := IParent;
       IParent := FormsCounterToParent(IChild);
    end;
end;

function TFMAIN.FormsCaptionToCounter(const SCaption, SYear, SMonth: String; const OnlyMainBank: Boolean): Integer;
var LForma_ : TLForma;
    Dat     : TDate;
begin
    {�������������}
    Result := -1;
    If (SYear = '') or (SMonth = '') then Exit;
    Dat:=SDatePeriod(SYear, SMonth);
    If Dat=0 then Exit;

    try     If Not FindForms(@LForma_, false, OnlyMainBank, -1, -1, '', SCaption, '', SMonth, false, Dat) then Exit;
            If Length(LForma_) <> 1 then Exit;
            Result := LForma_[0].FCounter;
    finally SetLength(LForma_, 0);
    end;
end;


{==============================================================================}
{=============  �������� ��������� ����������� � �������������� �����  ========}
{==============================================================================}
function TFMAIN.IsFormPeriod(const SCaption, SYear, SMonth: String; const Is1, Is3: Boolean): Boolean;
var LForma_ : TLForma;
    Dat     : TDate;
begin
    {�������������}
    Result := false;
    If (SCaption = '') or (SYear = '') or (SMonth = '') then Exit;
    Dat:=SDatePeriod(SYear, SMonth);
    If Dat=0 then Exit;

    {������ ��� �������� ����� ������}
    try     If Not FindForms(@LForma_, true, true, -1, -1, '', SCaption, '', '', false, Dat) then Exit;
            If Length(LForma_) <> 1 then Exit;
            Result := true;
            If Is1 then Result := Result and LForma_[0].FMonth1;
            If Is3 then Result := Result and (Not LForma_[0].FMonth1) and LForma_[0].FMonth3;
    finally SetLength(LForma_, 0);
    end;
end;


{==============================================================================}
{====  ��������� ������ ��������, ��������������� ������������� ��������  =====}
{==============================================================================}
{===================  P ����� ���� = nil                         ==============}
{===================  Result = true - ������� ���� ���� ������   ==============}
{==============================================================================}
function TFMAIN.FindRegions(const P        : PLRegion; const FindFirst : Boolean;
                            const FCounter : Integer;  const FCaption  : String;
                            const FParent  : Integer;  const FGroup    : Integer;
                            const FDate    : TDate): Boolean;
var BCounter, BCaption, BParent, BGroup, BDate : Boolean;
    I, ICount                                  : Integer;
begin
    {�������������}
    Result := false;
    If P<>nil then SetLength(P^, 0);
    BCounter := (FCounter >= 0);
    BCaption := (FCaption <> '');
    BParent  := (FParent  >= 0);
    BGroup   := (FGroup   >= 0);
    BDate    := (FDate    >  0);

    {������������� ��� ������}
    ICount:=Length(LRegion)-1;
    For I:=0 to ICount do begin
       If BCaption then If LRegion[I].FCaption <> FCaption then Continue;
       If BParent  then If LRegion[I].FParent  <> FParent  then Continue;
       If BCounter then If LRegion[I].FCounter <> FCounter then Continue;
       If BGroup   then If LRegion[I].FGroup   <> FGroup   then Continue;
       If BDate    then begin
          If LRegion[I].FBegin > 0 then If LRegion[I].FBegin > FDate then Continue;
          If LRegion[I].FEnd   > 0 then If LRegion[I].FEnd   < FDate then Continue;
       end;

       {������������ ���������}
       Result:=true;
       If P=nil then Break;

       {��������� ������}
       SetLength(P^, Length(P^)+1);
       P^[Length(P^)-1] := LRegion[I];

       {���� ���������� ������ ������ ������}
       If FindFirst then Break;
    end;
end;


{==============================================================================}
{======  ��������� ������ ������, ��������������� ������������� ��������  =====}
{==============================================================================}
{===================  P ����� ���� = nil                         ==============}
{===================  Result = true - ������� ���� ���� ������   ==============}
{==============================================================================}
function TFMAIN.FindTables(const P: PLTable; const FindFirst: Boolean; const FCounter, FCaption: String; const FDate: TDate): Boolean;
var BCounter, BCaption, BDate : Boolean;
    I, ICount                 : Integer;
begin
    {�������������}
    Result := false;
    If P<>nil then SetLength(P^, 0);
    BCounter := (FCounter <> '');
    BCaption := (FCaption <> '');
    BDate    := (FDate    >  0);

    {������������� ��� ������}
    ICount:=Length(LTable)-1;
    For I:=0 to ICount do begin
       If BCounter then If LTable[I].FCounter <> FCounter then Continue;
       If BCaption then If LTable[I].FCaption <> FCaption then Continue;
       If BDate    then begin
          If LTable[I].FBegin > 0 then If LTable[I].FBegin > FDate then Continue;
          If LTable[I].FEnd   > 0 then If LTable[I].FEnd   < FDate then Continue;
       end;

       {������������ ���������}
       Result:=true;
       If P=nil then Break;

       {��������� ������}
       SetLength(P^, Length(P^)+1);
       P^[Length(P^)-1] := LTable[I];

       {���� ���������� ������ ������ ������}
       If FindFirst then Break;
    end;
end;


{==============================================================================}
{======  ��������� ������ �����, ��������������� ������������� ��������  ======}
{==============================================================================}
{===================  P ����� ���� = nil                         ==============}
{===================  Result = true - ������� ���� ���� ������   ==============}
{==============================================================================}
function TFMAIN.FindRows(const P: PLRow;  const FindFirst: Boolean;
                         const FNumeric, FCaption: String; const FTable: Integer): Boolean;
var BNumeric, BCaption, BTable : Boolean;
    I, ICount                  : Integer;
begin
    {�������������}
    Result := false;
    If P<>nil then SetLength(P^, 0);
    BNumeric := (FNumeric <> '');
    BCaption := (FCaption <> '');
    BTable   := (FTable   >  0);

    {������������� ��� ������}
    ICount:=Length(LRow)-1;
    For I:=0 to ICount do begin
       If BTable   then If LRow[I].FTable   <> FTable   then Continue;
       If BNumeric then If LRow[I].FNumeric <> FNumeric then Continue;
       If BCaption then If LRow[I].FCaption <> FCaption then Continue;

       {������������ ���������}
       Result:=true;
       If P=nil then Break;

       {��������� ������}
       SetLength(P^, Length(P^)+1);
       P^[Length(P^)-1] := LRow[I];

       {���� ���������� ������ ������ ������}
       If FindFirst then Break;
    end;
end;


{==============================================================================}
{=====  ��������� ������ ��������, ��������������� ������������� ��������  ====}
{==============================================================================}
{===================  P ����� ���� = nil                         ==============}
{===================  Result = true - ������� ���� ���� ������   ==============}
{==============================================================================}
function TFMAIN.FindCols(const P: PLCol; const FindFirst: Boolean;
                         const FNumeric, FCaption: String; const FTable: Integer): Boolean;
var BNumeric, BCaption, BTable : Boolean;
    I, ICount                  : Integer;
begin
    {�������������}
    Result := false;
    If P<>nil then SetLength(P^, 0);
    BNumeric := (FNumeric <> '');
    BCaption := (FCaption <> '');
    BTable   := (FTable   >  0);

    {������������� ��� ������}
    ICount:=Length(LCol)-1;
    For I:=0 to ICount do begin
       If BTable   then If LCol[I].FTable   <> FTable   then Continue;
       If BNumeric then If LCol[I].FNumeric <> FNumeric then Continue;
       If BCaption then If LCol[I].FCaption <> FCaption then Continue;

       {������������ ���������}
       Result:=true;
       If P=nil then Break;

       {��������� ������}
       SetLength(P^, Length(P^)+1);
       P^[Length(P^)-1] := LCol[I];

       {���� ���������� ������ ������ ������}
       If FindFirst then Break;
    end;
end;


{==============================================================================}
{=======  ��������� ������ ����, ��������������� ������������� ��������  ======}
{==============================================================================}
{===================  P ����� ���� = nil                         ==============}
{===================  IsID = true - ������ ������ � ID           ==============}
{===================  Result = true - ������� ���� ���� ������   ==============}
{==============================================================================}
function TFMAIN.FindForms(const P: PLForma; const FindFirst, OnlyMainBank: Boolean; const FCounter, FParent: Integer;
                          const FRegion, FCaption, FFile, FMonth: String; const IsID: Boolean; const FDate: TDate): Boolean;
var BCounter, BParent, BRegion, BCaption, BFile, BMonth, BMonth1, BMonth3, BDate : Boolean;
    I, ICount : Integer;
begin
    {�������������}
    Result := false;
    If P<>nil then SetLength(P^, 0);
    BCounter := (FCounter >= 0);
    BParent  := (FParent  >= 0);
    BRegion  := (FRegion  <> '');
    BCaption := (FCaption <> '');
    BMonth   := (FMonth   <> '');
    BFile    := (FFile    <> '');
    BDate    := (FDate    >  0);
    If BMonth then begin
       BMonth1 := FindStrInArray(FMonth, ['������', '�������', '������',   '���',  '����', '������', '�������', '������']);
       BMonth3 := FindStrInArray(FMonth, ['����',   '����',    '��������', '�������']);
    end else begin
       BMonth1 := false;
       BMonth3 := false;
    end;

    {������������� ��� ������}
    ICount:=Length(LForma)-1;
    For I:=0 to ICount do begin
       If OnlyMainBank then If LForma[I].SBankPref <> ''       then Continue;
       If BCounter     then If LForma[I].FCounter  <> FCounter then Continue;
       If BParent      then If LForma[I].FParent   <> FParent  then Continue;
       If BRegion      then If (LForma[I].FRegion  <> FRegion) and (LForma[I].FRegion <> '') then Continue;
       If BCaption     then If LForma[I].FCaption  <> FCaption then Continue;
       If BMonth and BMonth1 then If LForma[I].FMonth1 = false then Continue;
       If BMonth and BMonth3 then If LForma[I].FMonth3 = false then Continue;
       If BFile        then If LForma[I].FFile     <> FFile    then Continue;
       If BDate        then begin
          If LForma[I].FBegin > 0 then If LForma[I].FBegin > FDate then Continue;
          If LForma[I].FEnd   > 0 then If LForma[I].FEnd   < FDate then Continue;
       end;
       If IsID then begin
          If (LForma[I].FCellID = '') or (LForma[I].FStrID = '') then Continue;
       end;


       {������������ ���������}
       Result:=true;
       If P=nil then Break;

       {��������� ������}
       SetLength(P^, Length(P^)+1);
       P^[Length(P^)-1] := LForma[I];

       {���� ���������� ������ ������ ������}
       If FindFirst then Break;
    end;
end;


{==============================================================================}
{=====  ��������� ������ ������, ��������������� ������������� ��������  ======}
{==============================================================================}
{===================  P ����� ���� = nil                         ==============}
{===================  Result = true - ������� ���� ���� ������   ==============}
{==============================================================================}
function TFMAIN.FindORKeys(const P: PLORKey; const FindFirst: Boolean; const FKey1, FKey2: String): Boolean;
var BKey1, BKey2 : Boolean;
    I, ICount    : Integer;
begin
    {�������������}
    Result := false;
    If P<>nil then SetLength(P^, 0);
    BKey1 := (FKey1 <> '');
    BKey2 := (FKey2 <> '');

    {������������� ��� ������}
    ICount:=Length(LORKey)-1;
    For I:=0 to ICount do begin
       If BKey1 then If LORKey[I].FKey1 <> FKey1 then Continue;
       If BKey2 then If LORKey[I].FKey2 <> FKey2 then Continue;

       {������������ ���������}
       Result:=true;
       If P=nil then Break;

       {��������� ������}
       SetLength(P^, Length(P^)+1);
       P^[Length(P^)-1] := LORKey[I];

       {���� ���������� ������ ������ ������}
       If FindFirst then Break;
    end;
end;


{==============================================================================}
{===========================   ���� ��������� ���   ===========================}
{==============================================================================}
function TFMAIN.Finder(const Directly: Boolean): Boolean;
label Nx;
type TInd = record
        ITab, IRow, ICol: Integer;
     end;
var FBASE         : TBASE;
    SList, LFind  : TStringList;
    SFind, S, YMR : String;
    LTab_         : TLTable;
    LRow_         : TLRow;
    LCol_         : TLCol;
    LForma_       : TLForma;
    Ind, Ind0     : TInd;
    Val           : Extended;
    I             : Integer;

    {��������� ������� ������� Ind0}
    function GetFirstInd: Boolean;
    var NForm            : TTreeNode;
        STab, SRow, SCol : String;
        I                : Integer;
    begin
        {�������������}
        Result:=false;

        {��� ����������}
        STab := '1'; SRow := '0'; SCol := '0';
        NForm := TreeForm.Selected;
        If NForm<>nil then begin
           If NForm.ImageIndex = ICO_FORM_DETAIL_COL then begin
              STab := IntToStr(Integer(NForm.Parent.Parent.Data));
              SRow := IntToStr(Integer(NForm.Parent.Data));
              SCol := IntToStr(Integer(NForm.Data));
           end;
           If NForm.ImageIndex = ICO_FORM_DETAIL_ROW then begin
              STab := IntToStr(Integer(NForm.Parent.Data));
              SRow := IntToStr(Integer(NForm.Data));
           end;
           If NForm.ImageIndex = ICO_FORM_DETAIL_TABLE then begin
              STab := IntToStr(Integer(NForm.Data));
           end;
        end;

        {���������� ��������� ������ �������}
        If Length(LTable)=0 then Exit;
        Ind0.ITab := Low(LTable);
        For I:=Low(LTable) to High(LTable) do begin
           If LTable[I].FCounter = STab then begin
              Ind0.ITab := I;
              Break;
           end;
        end;

        {���������� ������ ����� � ��������� ������ ������}
        If Not FindRows(@LRow_, false, '',   '', StrToInt(STab)) then Exit;
        Ind0.IRow := Low(LRow_);
        For I:=Low(LRow_) to High(LRow_) do begin
           If LRow_[I].FNumeric = SRow then begin
              Ind0.IRow := I;
              Break;
           end;
        end;

        {���������� ������ �������� � ��������� ������ �������}
        If Not FindCols(@LCol_, false, '',   '', StrToInt(STab)) then Exit;
        Ind0.ICol := Low(LCol_);
        For I:=Low(LCol_) to High(LCol_) do begin
           If LCol_[I].FNumeric = SCol then begin
              Ind0.ICol := I;
              Break;
           end;
        end;

        {������������ ���������}
        Result:=true;
    end;

    {����� � ������ ������ ������� ��������� ������ �������}
    function NextTab: Boolean;
    begin
        {�������������}
        Result:=false;

        {������������ ������ �������}
        If Directly then begin
           Ind.ITab := Ind.ITab + 1;
           If Ind.ITab > High(LTable) then Ind.ITab := Low(LTable);
        end else begin
           Ind.ITab := Ind.ITab - 1;
           If Ind.ITab < Low(LTable) then Ind.ITab := High(LTable);
        end;

        {������������ ������ ����� � ��������}
        If Not FindCols(@LCol_, false, '', '', StrToInt(LTable[Ind.ITab].FCounter)) then Exit;
        If Not FindRows(@LRow_, false, '', '', StrToInt(LTable[Ind.ITab].FCounter)) then Exit;
        If Directly then begin
           Ind.ICol := Low(LCol_);
           Ind.IRow := Low(LRow_);
        end else begin
           Ind.ICol := High(LCol_);
           Ind.IRow := High(LRow_);
        end;

        {������������ ���������}
        Result:=true;
    end;

    {����� � ������ ����� ������� ��������� ������ ������}
    function NextRow: Boolean;
    begin
        If Directly then begin
           Ind.ICol := Low(LCol_);
           Ind.IRow := Ind.IRow + 1;
        end else begin
           Ind.ICol := High(LCol_);
           Ind.IRow := Ind.IRow - 1;
        end;
        If (Ind.IRow < Low(LRow_)) or (Ind.IRow > High(LRow_)) then Result := NextTab
                                                               else Result := true;
    end;

    {����� � ������ �������� ������� ��������� ������ �������}
    function NextCol: Boolean;
    begin
        If Directly then Ind.ICol := Ind.ICol + 1
                    else Ind.ICol := Ind.ICol - 1;
        If (Ind.ICol < Low(LCol_)) or (Ind.ICol > High(LCol_)) then Result := NextRow
                                                               else Result := true;
    end;

    {����� ����� LFind � ������ Str}
    function FindTextForm(const Str: String): Boolean;
    var I: Integer;
    begin
        {�������������}
        Result:=false;

        {���� ��� ���������}
        For I:=0 to LFind.Count-1 do begin
           If Pos(LFind[I], Str)<=0 then begin
              Exit;
           end;
        end;

        {������������ ���������}
        Result:=true;
    end;

begin
    {�������������}
    Result := false;
    SFind  := AnsiUpperCase(Trim(CBFind.Text));
    If SFind = '' then Exit;

    {������������ � ��������� ������ ������}
    S := Trim(CBFind.Text);
    I := CBFind.Items.IndexOf(S);
    If I>=0 then CBFind.Items.Delete(I);
    CBFind.Items.Insert(0, S);
    CBFind.Text := S;
    For I:=CBFind.Items.Count-1 downto 25 do CBFind.Items.Delete(I);
    WriteCBListIni(@CBFind, INI_FIND);

    If Not SetSelPath then Exit;
    YMR := SetYMR(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '');

    SetLength(LTab_, 0);
    SetLength(LRow_, 0);
    SetLength(LCol_, 0);
    SList := TStringList.Create;
    LFind := TStringList.Create;
    FBASE := TBASE.Create;
    try
       {��������� ������ ������}
       While SFind<>'' do LFind.Add(Trim(TokChar(SFind, ' ')));
       If LFind.Count = 0 then Exit;

       {��������� ������� ������� Ind0}
       If Not GetFirstInd then NextTab;

       {��������� ������� ����� �������}
       Ind:=Ind0;

 Nx:   {��������� ��� � ������ Direct}
       If Not NextCol then Exit;

       {������ ������ ���� � ������ �� �����}
       If  (Ind.ITab = Ind0.ITab)
       and (Ind.IRow = Ind0.IRow)
       and (Ind.ICol = Ind0.ICol) then begin Beep; Exit; end;

       {����}
       S:=AnsiUpperCase(LTable[Ind.ITab].FCaption+'|'+LRow_[Ind.IRow].FCaption+'|'+LCol_[Ind.ICol].FCaption);
       If FindTextForm(S) then begin
          {��������� ��� ������ ����� ��������}
          Val := FBASE.GetValYMR(YMR, LTable[Ind.ITab].FCounter, LCol_[Ind.ICol].FNumeric, LRow_[Ind.IRow].FNumeric);
          If Val > 0 then begin
             {���� �������������� �����, ������� ����������� ���}
             If Not FindForms(@LForma_, false, false, -1, -1, '', '', '', SelPath.SMonth, false, SelPath.Date) then Goto Nx;
             For I:=Low(LForma_) to High(LForma_) do begin
                If LForma_[I].FFile = '' then Continue;
                If LForma_[I].FIcon <> (ICO_FORM_MIN - ICO_FORM_DETAIL_REPORT + 1) then Continue;
                S := PathFileForm(LForma_[I].SBankPref, LForma_[I].FFile);
                S := ChangeFileExt(S, '.ini');
                If Not GetListTabIni(@SList , S) then Continue;
                {���� ����� �������}
                If SList.IndexOf(LTable[Ind.ITab].FCounter)>=0 then begin
                   S:=FormsCounterToPathTreeCaption(LForma_[I].FCounter)+
                      LTable[Ind.ITab].FCaption+'\'+
                      LRow_[Ind.IRow].FCaption+'\'+
                      LCol_[Ind.ICol].FCaption+' = '+FloatToStr(Val)+'\';
                   WriteLocalString(INI_SELECT, INI_SELECT_FORM, S);
                   RefreshTreeForm;
                   Exit;
                end;
             end;
          end;
       end;
       Goto Nx;

    finally
       FBASE.Free;
       LFind.Free; 
       SList.Free;
       SetLength(LForma_, 0);
       SetLength(LTab_,   0);
       SetLength(LRow_,   0);
       SetLength(LCol_,   0);
    end;
end;


