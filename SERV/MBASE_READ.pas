{==============================================================================}
{==========================   ������ ��������   ===============================}
{==============================================================================}
function TBASE.GetValYMR(const YMR, STab, SCol, SRow: String): Extended;
var SPath : String;
begin
    SPath  := FFMAIN.TabToDatYMR(STab, YMR);
    Result := GetVal(SPath, STab, SCol, SRow);
end;


{==============================================================================}
{==========================   ������ ��������   ===============================}
{==============================================================================}
function TBASE.GetVal(const SPath, STab, SCol, SRow: String): Extended;
var FIni                            : TMemIniFile;
    SPath0                          : String;
    SKey, SKeyTab, SKeyCol, SKeyRow : String;
    LORKey_                         : TLORKey;
    S                               : String;
    I                               : Integer;

    SYear0, SMonth0, SRegion0       : String;
    SYear,  SMonth                  : String;
    DMonth                          : Integer;
begin
    {�������������}
    Result := 0;
    SPath0 := SPath;               // D:\���� ������\2009\����\�������������.dat

    {*** ���� �������� �������� ���������-������������� ����������� ***********}
    If FFMAIN.LCEP.IndexOf(STab) >= 0 then begin
       {*** ����������� �������� (��������� ������) ***************************}
       If BSEPDetail then begin
          {�������������}
          SKey     := SCol+' '+SRow;
          SRegion0 := TokCharEnd(SPath0, '\');
          SMonth0  := TokCharEnd(SPath0, '\');
          SYear0   := TokCharEnd(SPath0, '\');
          {������������� �������}
          For DMonth:=0 to ISEPMonth do begin
             {������������ ������}
             SYear  := SYear0;
             SMonth := SMonth0;
             If Not CorrectPeriodStr(SYear, SMonth, DMonth*(-1)) then Break;
             {������ ��������}
             FIni:=TMemIniFile.Create(SPath0+'\'+SYear+'\'+SMonth+'\'+SRegion0);
             try     Result:=FIni.ReadInteger(STab, SKey, 0);
             finally FIni.Free;
             end;
             {��������� ��������}
             If Result <> 0 then Break;
          end;
          {�����}
          Exit;
       {*** ����������� ��������� *********************************************}
       end else begin
          {������� ������ �������� �� ������ �������� ����}
          SPath0 := ReplSlovoEndChar(SPath0, '������', 2, '\');
          GetCash(SPath0, STab, SCol, SRow, Result);
          If Result <> 0 then Exit;
          {������� ������ �������� �� ������ ����������� ����}
          SYear0 := CutSlovoEndChar(SPath0, 3, '\');
          try SYear0 := IntToStr(StrToInt(SYear0)-1); finally end;
          SPath0 := ReplSlovoEndChar(SPath0, SYear0, 3, '\');
       end;
    end;

    {*** ������ ������ �� ���� ************************************************}
    If GetCash(SPath0, STab, SCol, SRow, Result) then If Result <> 0 then Exit;                       // If Result <> 0 ����� ��������� ������

    {*** �������������� ������ ������ *****************************************}
    S := STab+' '+SCol+' '+SRow;
    try
       {�������� �����������}
       If FFMAIN.FindORKeys(@LORKey_, false, S, '') then begin
          For I:=Low(LORKey_) to High(LORKey_) do begin
             SKey    := LORKey_[I].FKey2;
             SKeyTab := TokChar(SKey, ' ');
             SKeyCol := TokChar(SKey, ' ');
             SKeyRow := TokChar(SKey, ' ');
             If GetCash(SPath, SKeyTab, SKeyCol, SKeyRow, Result) then If Result <> 0 then Exit;     // If Result <> 0 ����� ��������� ������
          end;
       end;
       If FFMAIN.FindORKeys(@LORKey_, false, '', S) then begin
          For I:=Low(LORKey_) to High(LORKey_) do begin
             SKey    := LORKey_[I].FKey1;
             SKeyTab := TokChar(SKey, ' ');
             SKeyCol := TokChar(SKey, ' ');
             SKeyRow := TokChar(SKey, ' ');
             If GetCash(SPath, SKeyTab, SKeyCol, SKeyRow, Result) then If Result <> 0 then Exit;    // If Result <> 0 ����� ��������� ������
          end;
       end;
    finally
       SetLength(LORKey_, 0);
    end;
end;


{==============================================================================}
{===================       ����� DAT-����� � ����       =======================}
{==============================================================================}
{===================   SPath  - ���� � �������� �����   =======================}
{===================   RESULT - ������ ����             =======================}
{==============================================================================}
function TBASE.FindCashDat(const SPath: String): Integer;
var I      : Integer;
    SPath0 : String;
begin
    Result := -1;
    SPath0 := AnsiUpperCase(SPath);
    For I:=Low(CashTab) to High(CashTab) do begin
       If AnsiUpperCase(CashTab[I].SPath) = SPath0 then begin
          Result := I;
          Break;
       end;
    end;
end;


{==============================================================================}
{===================       ����� ������� � ����         =======================}
{==============================================================================}
{===================   STab   - ������� �������         =======================}
{===================   RESULT - ������ ����             =======================}
{==============================================================================}
function TBASE.FindCashTab(const SPath, STab: String): Integer;
var I      : Integer;
    SPath0 : String;
begin
    Result := -1;
    SPath0 := AnsiUpperCase(SPath);
    For I:=Low(CashTab) to High(CashTab) do begin
       If (AnsiUpperCase(CashTab[I].SPath) = SPath0) and (CashTab[I].STab = STab) then begin
          Result := I;
          Break;
       end;
    end;
end;


{==============================================================================}
{===================   �������� DAT-����� � ���         =======================}
{==============================================================================}
function TBASE.LoadCashDat(const SPath: String): Boolean;
var LMem           : TStringList;
    I, IMem        : Integer;
    SVal, STabCash : String;
begin
    {�������������}
    Result := false;

    {���� ����� ������������ dat-����}
    Result := (FindCashDat(SPath) >=0 );
    If Result then Exit;

    {��� ���������� ����������� ��������� dat-���� ���������}
    If Not FileExists(SPath) then Exit;
    I    := 0;
    LMem := TStringList.Create;
    try
       LMem.LoadFromFile(SPath);
       For IMem:=0 to LMem.Count-1 do begin
          SVal := LMem[IMem];
          STabCash := CutModulChar(SVal, '[', ']');
          If STabCash <> '' then begin
             {�������� ��������� �������}
             I := Length(CashTab);
             SetLength(CashTab, I+1);
             CashTab[I].SPath := SPath;
             CashTab[I].STab  := STabCash;
             CashTab[I].SList := TStringList.Create;
             Result := true;
          end else begin
             {�������� �������� �������}
             CashTab[I].SList.Add(SVal);
          end;
       end;
    finally
       LMem.Free;
    end;
end;


{==============================================================================}
{========================   ������ �������� �� ����   =========================}
{==============================================================================}
function TBASE.GetCash(const SPath, STab, SCol, SRow: String; var IVal: Extended): Boolean;
var S : String;
    I : Integer;
begin
    {�������������}
    Result := false;
    IVal   := 0;

    {���� ��� �������}
    I := FindCashTab(SPath, STab);

    {��� �� ���������� � ���� �������� ���������� ���� DAT-����}
    If I < 0 then begin
       If Not LoadCashDat(SPath) then Exit;
       {����� ����������� �������� ���� STab}
       I := FindCashTab(SPath, STab);
       If I < 0 then Exit;
    end;

    {������ ��������}
    S := CashTab[I].SList.Values[SCol+' '+SRow];
    If IsFloatStr(S) then IVal := StrToFloat(S);

    {������������ ���������}
    Result := true;
end;


