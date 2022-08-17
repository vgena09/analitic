{==============================================================================}
{===========    !!! ����� !!! ����� !!! ����� !!! ����� !!!     ===============}
{==============================================================================}
{===========         ��������� YMR � ����������� ������         ===============}
{==============================================================================}
{===========            YMR - Year\Month\Region.dat             ===============}
{==============================================================================}


{==============================================================================}
{=======   ����� ����� � ����� DAT-������ � ������ ������    ==================}
{==============================================================================}
{=======  SFindYMR - Year\Month\* ��� Year\Month\Region.dat  ==================}
{=======  PSr^     - ������ ������ ����� (�/�=nil)           ==================}
{=======  IsFullPath = true  - � PSr ������ ����             ==================}
{=======  IsFullPath = false - � PSr ����������� �������     ==================}
{=======  Result             - ����� ������������ ������     ==================}
{==============================================================================}
function TFMAIN.FindDataYMR(const PSr: PArrayStr; const SFindYMR: String; const IsFullPath: Boolean): Integer;
var Sr           : TSearchRec;
    IBank        : Integer;
    SName, SPath : String;

    function IsNoRecord: Boolean;
    begin
        Result := true;
        If IsFullPath or (PSr = nil) then Exit;
        Result := Not FindStrInArray(SName, PSr^);
    end;
begin
    {�������������}
    Result := 0;
    If PSr <> nil then SetLength(PSr^, 0);
    {������������� ����� ������}
    For IBank := Low(LBank) to High(LBank) do begin
       SPath := PrefToDatYMR(LBank[IBank].Pref, SFindYMR);
       {����� �����}
       try
          If FindFirst(SPath, faDirectory, Sr) = 0 then begin
             SPath := ExtractFilePath(SPath);
             repeat
                SName := Sr.Name;
                If IsNoRecord and (SName <> '.') and (SName <> '..') then begin
                   Result := Result + 1;
                   If PSr <> nil then begin
                      SetLength(PSr^, Result);
                      If IsFullPath then PSr^[Result-1] := SPath + SName
                                    else PSr^[Result-1] := SName;
                   end;
                end;
             until FindNext(Sr) <> 0;
          end;
       finally
          FindClose(Sr);
       end;
    end;
end;


{==============================================================================}
{===  ������� � DAT-����� ���� ����� �������, ����������������� � INI-����� ===}
{==============================================================================}
function TFMAIN.IsTableDatIncludeIniYMR(const SPathIni, YMR: String): Boolean;
var LIni, LDat : TStringList;
    I          : Integer;
begin
    {�������������}
    Result := false;
    LIni   := TStringList.Create;
    LDat   := TStringList.Create;
    try
       {������ ������}
       If Not GetListTabIni   (@LIni, SPathIni) then Exit;
       If Not GetListTabDatYMR(@LDat, YMR) then Exit;

       {��������}
       For I:=0 to LDat.Count-1 do begin
          If LIni.IndexOf(LDat[I])>=0 then begin
             Result:=true;
             Break;
          end;
       end;
    finally
       LDat.Free;
       LIni.Free;
    end;
end;


{==============================================================================}
{===========        ��������� YMR � ������ ��������� SKEY         =============}
{==============================================================================}
{=====   SKEY =  ID ������� /                                             =====}
{=====           ID ������� /                                             =====}
{=====           ID ������  /                                             =====}
{=====           ��������� ������� � ������� / 0-�� ���������             =====}
{=====           ID ������� / 0-�� ��������� ��� �������� �������         =====}
{=====   SKEY = '' - ��������� �����������                                =====}
{==============================================================================}
{=====   ���� ID ������� ����� � ���� ����� - ��������� ������            =====}
{==============================================================================}
function TFMAIN.SetYMR(const SYear, SMonth, SRegion, SKey: String): String;
var SYear0, SMonth0, SRegion0, SKey0  : String;
    SKeyTable, SKeyPeriod, SKeyRegion : String;
    IKeyPeriod                        : Integer;
    IYear, IMonth                     : Integer;
begin
    {�������������}
    Result     := '';
    SYear0     := SYear;
    SMonth0    := SMonth;
    SRegion0   := SRegion;

    {��� ������������� ���������� ���������}
    If SKey<>'' then begin
       SKey0      := SKey;
       SKeyTable  := CutSlovo(SKey0, 1, CH_SPR);
       SKeyPeriod := CutSlovo(SKey0, 4, CH_SPR);
       SKeyRegion := CutSlovo(SKey0, 5, CH_SPR);

       {������������ ������}
       If (SKeyPeriod<>'') and (SKeyPeriod<>'0') then begin
          IKeyPeriod := StrToInt(SKeyPeriod);
          IYear      := StrToInt(SYear0);
          IMonth     := MonthStrToInd(SMonth0);

          IYear      := IYear  + (IKeyPeriod Div 12);
          IMonth     := IMonth + (IKeyPeriod Mod 12);

          If IMonth > High(MonthList) then begin Inc(IYear); IMonth:=IMonth-12; end;
          If IMonth < Low(MonthList)  then begin Dec(IYear); IMonth:=IMonth+12; end;

          SYear0     := IntToStr(IYear);
          SMonth0    := MonthIndToStr(IMonth);
       end;

       {������������ ������}
       If (SKeyRegion<>'') and (SKeyRegion<>'0') then begin
          If IsIntegerStr(SKeyRegion) then SRegion0 := RegionsCounterToCaption(StrToInt(SKeyRegion))
                                      else SRegion0 := SKeyRegion;
       end;
    end;

    {���������� ����}
    Result := SYear0+'\'+SMonth0+'\'+SRegion0+'.dat';
end;


{==============================================================================}
{===========   ������� �� ���� ���� ������ ��� �������/������   ===============}
{==============================================================================}
{===========                SROW = '' - ����� ��������          ===============}
{==============================================================================}
function TFMAIN.IsExistsRowYMR(const YMR, STab, SRow: String): Boolean;
begin Result := IsExistsRow(TabToDatYMR(STab, YMR), STab, SRow);
end;

{==============================================================================}
{===================     ��������� ���� � DAT-�����     =======================}
{==============================================================================}
function TFMAIN.TabToDatYMR(const STab, YMR: String): String;
begin Result := PrefToDatYMR(TablesCounterToBankPrefix(STab, YMRToDat(YMR)), YMR);
end;
function TFMAIN.PrefToDatYMR(const SPref, YMR: String): String;
begin Result := PATH_BD0+SPref+'\'+PATH_DATA+YMR;
end;


{==============================================================================}
{===============   ������ ����� ������ �� ����� DAT-������   ==================}
{==============================================================================}
function TFMAIN.GetTabKeyYMR(const PList: PStringList; const YMR, Section: String): Boolean;
begin
    Result := GetTabKey(PList, PATH_BD0+TablesCounterToBankPrefix(Section, YMRToDat(YMR))+'\'+PATH_DATA+YMR, Section);
end;


{==============================================================================}
{=============   �������� ������ ������ �� ����� DAT-������   =================}
{==============================================================================}
function TFMAIN.GetListTabDatYMR(const PList: PStringList; const YMR: String): Boolean;
var LFind : array of String;
    SList : TStringList;
    IFind : Integer;
begin
    {�������������}
    Result:=false;
    If PList=nil then Exit;
    PList^.Clear;
    SList := TStringList.Create;
    try
       {�������� � ������������� ������ dat-������}
       If FindDataYMR(@LFind, YMR, true) = 0 then Exit;
       For IFind := Low(LFind) to High(LFind) do begin
          {������ ������ ������ � SList}
          GetListTabDat(@SList, LFind[IFind]);
          {���������� ������ ������ �������� ����������}
          SumSListUniq(PList, @SList);
       end;
    finally
       SetLength(LFind, 0);
       SList.Free;
    end;
    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{========================   ����������� YMR � DATE   ==========================}
{==============================================================================}
function TFMAIN.YMRToDat(const YMR: String): TDate;
var SYear, SMonth : String;
begin
    SYear  := CutSlovo(YMR, 1, '\');
    SMonth := CutSlovo(YMR, 2, '\');
    Result := SDatePeriod(SYear, SMonth);
end;

(*
{==============================================================================}
{=======  ������������ ���� � DAT-����� � ����������� �� �������   ============}
{==============================================================================}
{=======   �� ��������� SPATH - DAT-���� ��������� ����� ������    ============}
{==============================================================================}
function TFMAIN.CorrectPathDat(const SPath, STab: String): String;
var LTable_ : TLTable;
begin
    Result := SPath;
    try
       If Not FindTables(@LTable_, true, STab, '', 0) then Exit;
       If LTable_[0].SBankPref = '' then Exit;
       Delete(Result, 1, Length(PATH_BD0));
       Result := PATH_BD0 + LTable_[0].SBankPref + Result;
    finally
       SetLength(LTable_, 0);
    end;
end;
*)

