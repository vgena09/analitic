{==============================================================================}
{===========             ������ � �������� (���������)          ===============}
{==============================================================================}
{===========    !!! ����� !!! ����� !!! ����� !!! ����� !!!     ===============}
{==============================================================================}
{===========  ������ � ����������� ������, YMR �� ������������  ===============}
{==============================================================================}


{==============================================================================}
{=================   �������� ������ ������ �� INI-�����   ====================}
{==============================================================================}
function TFMAIN.GetListTabIni(const PList: PStringList; const SPathIni: String): Boolean;
var FIni : TIniFile;
    S    : String;
begin
    {�������������}
    Result:=false;
    If PList=nil then Exit;
    PList^.Clear;
    If Not FileExists(SPathIni) then Exit;

    {������ ������}
    FIni:=TIniFile.Create(SPathIni);
    try     S := FIni.ReadString(INI_PARAM, INI_PARAM_TABLES, '');
            If Not SeparatMStr(S, PList, ', ') then Exit;
    finally FIni.Free;
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{=================   �������� ������ ������ �� DAT-�����   ====================}
{==============================================================================}
function TFMAIN.GetListTabDat(const PList: PStringList; const SPathDat: String): Boolean;
var FIni : TMemIniFile;
begin
    {�������������}
    Result:=false;
    If PList=nil then Exit;
    PList^.Clear;
    If Not FileExists(SPathDat) then Exit;

    {������ ������}
    FIni:=TMemIniFile.Create(SPathDat);
    try     FIni.ReadSections(PList^);
    finally FIni.Free;
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{===================   ������ ����� ������ �� DAT-�����   =====================}
{==============================================================================}
function TFMAIN.GetTabKey(const PList: PStringList; const SPathDat, Section: String): Boolean;
begin
    {�������������}
    Result:=false;
    If PList=nil then Exit;
    PList^.Clear;
    If Not FileExists(SPathDat) then Exit;

    {������ ������}
    Result := (ReadSectionIni(PList, SPathDat, Section) >= 0);
    If Not Result then PList^.Clear;
end;


{==============================================================================}
{============================   ������ ������   ===============================}
{==============================================================================}
function TFMAIN.GetTabVal(const SPath, STab: String; const PList: PStringList): Boolean;
var FDat: TMemIniFile;
begin
    {�������������}
    Result:=false;
    If PList=nil then Exit;

    {������ ������}
    PList^.Clear;
    FDat:=TMemIniFile.Create(SPath);
    try     FDat.ReadSectionValues(STab, PList^);
    finally FDat.Free;
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{==================           ���������� ������             ===================}
{==============================================================================}
{==================   IsFullReplace - ������ ������ ������  ===================}
{==============================================================================}
function TFMAIN.SetTabVal(const SPath, STab: String; const PList: PStringList; const IsFullReplace: Boolean): Boolean;
var FDat    : TMemIniFile;
    S, SVal : String;
    I       : Integer;
begin
    {�������������}
    Result:=false;

    {������� ��������}
    S := ExtractFilePath(SPath);
    If Not DirectoryExists(S) then If Not ForceDirectories(S) then Exit;

    FDat:=TMemIniFile.Create(SPath);
    try
       {��� ������ ������ ������� c����� ���������}
       If IsFullReplace then FDat.EraseSection(STab);

       {���������� ������}
       For I:=0 to PList^.Count-1 do begin
          S    := PList^.Names[I];
          SVal := PList^.Values[S];
          If (SVal='') or (SVal='0') then FDat.DeleteKey(STab, S)
                                     else FDat.WriteString(STab, S, SVal);
       end;

       {������ ������ ������������}
       If Not FDat.SectionExists(STab) then begin
          FDat.WriteString(STab, '0', '0');
          FDat.DeleteKey(STab, '0');
       end;

       FDat.UpdateFile;
    finally
       FDat.Free;
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{======================   ������� ������ DAT-�����   ==========================}
{==============================================================================}
function TFMAIN.CreateTab(const SPath, STab: String): Boolean;
var FDat : TMemIniFile;
begin
    {�������������}
    Result:=false;

    {������� ������ ������}
    FDat:=TMemIniFile.Create(SPath);
    try     FDat.WriteString(STab, '0', '0');
            FDat.DeleteKey(STab, '0');
            FDat.UpdateFile;
    finally FDat.Free;
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{======================   ������� ������ DAT-�����   ==========================}
{==============================================================================}
function TFMAIN.ClearTab(const SPath, STab: String): Boolean;
begin
    {�������������}
    Result:=false;

    {������� ������}
    DelTab(SPath, STab);

    {������� ������ ������}
    CreateTab(SPath, STab);

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{======================   ������� ������ DAT-�����   ==========================}
{==============================================================================}
function TFMAIN.DelTab(const SPath, STab: String): Boolean;
var FDat : TMemIniFile;
begin
    {�������������}
    Result:=false;
    If Not FileExists(SPath) then begin Result:=true; Exit; end;

    {������� ������}
    FDat:=TMemIniFile.Create(SPath);
    try     FDat.EraseSection(STab);
            FDat.UpdateFile;
    finally FDat.Free;
    end;

    {������������ ���������}
    Result:=true;
end;


