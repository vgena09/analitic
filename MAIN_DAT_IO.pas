{==============================================================================}
{===========    !!! ����� !!! ����� !!! ����� !!! ����� !!!     ===============}
{==============================================================================}
{===========  ������ � ����������� ������, YMR �� ������������  ===============}
{==============================================================================}

{$INCLUDE MAIN_DAT_IO_SECTION}

{==============================================================================}
{=============   ������ �������� �� DAT-����� ��� �����������   ===============}
{==============================================================================}
function TFMAIN.GetValNoCash(const SPathDat, STab, SCol, SRow: String): String;
var FDat: TIniFile;
begin
    {�������������}
    Result:='';
    If Not FileExists(SPathDat) then Exit;

    {������ ������}
    FDat:=TIniFile.Create(SPathDat);
    try      Result:=FDat.ReadString(STab, SCol+' '+SRow, '');
    finally FDat.Free;
    end;
end;


{==============================================================================}
{====================   ���������� �� ������ DAT-�����   ======================}
{==============================================================================}
function TFMAIN.IsExistsTab(const SPath, STab: String): Boolean;
var LTab : TStringList;
begin
    {�������������}
    Result := false;

    {��������}
    LTab := TStringList.Create;
    try     If Not GetListTabDat(@LTab, SPath) then Exit;
            Result := (LTab.IndexOf(STab) >= 0);
    finally LTab.Free;
    end;
    // Result:=FDat.SectionExists(STab);   // �� ���������� ������ ������
end;


{==============================================================================}
{===================   ������� �� �������� �������� ������  ===================}
{==============================================================================}
{===  SRow = '' - ����� �������� (���� �� �����-���� �������� ��� �������) ====}
{==============================================================================}
function TFMAIN.IsExistsRow(const SPath, STab, SRow: String): Boolean;
var SList : TStringList;
    I     : Integer;
begin
    {�������������}
    Result := false;
    If Not FileExists(SPath) then Exit;

    SList := TStringList.Create;
    try
        {������ ������}
        If Not GetTabVal(SPath, STab, @SList) then Exit;
        {������������� ������}
        If SRow='' then begin
           Result := (SList.Count > 0);
        end else begin
           For I:=0 to SList.Count-1 do begin
              If CutSlovo(SList[I], 2, ' ') = SRow then begin
                 Result:=true;
                 Break;
              end;
           end;
        end;
    finally
       SList.Free;
    end;
end;



{==============================================================================}
{===  ������� � DAT-����� ���� ����� �������, ����������������� � INI-����� ===}
{==============================================================================}
function TFMAIN.IsTableDatIncludeIni(const SPathIni, SPathDat: String): Boolean;
var LIni, LDat : TStringList;
    I          : Integer;
begin
    {�������������}
    Result:=false;
    If Not FileExists(SPathDat) then Exit;
    LIni := TStringList.Create;
    LDat := TStringList.Create;
    try
       {������ ������}
       If Not GetListTabIni(@LIni, SPathIni) then begin ErrMsg('������ ������ ������ ������ �� INI-�����!'+CH_NEW+SPathIni); Exit; end;
       If Not GetListTabDat(@LDat, SPathDat) then begin ErrMsg('������ ������ ������ ������ �� DAT-�����!'+CH_NEW+SPathDat); Exit; end;

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
{============               ������ �� ����� ��� ������             ============}
{============      (�� �������� � ������� ����� ����� INI ����)    ============}
{==============================================================================}
function TFMAIN.IsFormReadOnly(const IDForm: Integer): Boolean;
var FIni                : TMemIniFile;
    LPage               : TStringList;
    SPathForm, SPathIni : String;
    S                   : String;
    I                   : Integer;
begin
    {�������������}
    Result := (FormsCounterToBankPrefix(IDForm) <> '');
    If Result then Exit else Result := true;

    SPathForm := FormsCounterToFile(IDForm);
    If Not FileExists(SPathForm) then Exit;
    SPathIni := ChangeFileExt(SPathForm, '.ini');
    If Not FileExists(SPathIni) then Exit;

    FIni  := TMemIniFile.Create(SPathIni);
    LPage := TStringList.Create;
    try
       {������ ������ ������}
       S:=FIni.ReadString(INI_PARAM, INI_PARAM_PAGES, '');
       If Not SeparatMStr(S, @LPage, ', ') then Exit;

       {������������� ������ ���� �� ����������� ����� ������}
       For I:=0 to LPage.Count-1 do begin
          S:=FIni.ReadString(INI_PAGE+LPage[I], INI_PAGE_IN, '');
          If S<>'' then begin
             Result:=false;
             Break;
          end;
       end;
    finally
       LPage.Free;
       FIni.Free;
    end;
end;


