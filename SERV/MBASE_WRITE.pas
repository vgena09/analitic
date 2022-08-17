{==============================================================================}
{========================   ���������� ��������   =============================}
{==============================================================================}
function TBASE.SetValYMR(const YMR, STab, SCol, SRow, SVal: String): Boolean;
var SPath : String;
begin
    SPath  := FFMAIN.TabToDatYMR(STab, YMR);
    Result := SetVal(SPath, STab, SCol, SRow, SVal);
end;


{==============================================================================}
{========================   ���������� ��������   =============================}
{==============================================================================}
function TBASE.SetVal(const SPath, STab, SCol, SRow, SVal: String): Boolean;
var SVal0 : String;
    PFile : PMemIniFile;
begin
    {�������������}
    Result := false;
    SVal0  := SVal;

    {���������� ���� � ������}
    PFile := FindMemFile(SPath);
    If PFile=nil then begin
       PFile := AddMemFile(SPath);
       If PFile=nil then begin ErrMsg('������ ����������� � dat-�����:'+CH_NEW+SPath); Exit; end;
    end;

    try
       {���� �����, �� ��������� �� ������}
       If (Pos(',', SVal0)>0) or (Pos('.', SVal0)>0) then begin
          SVal0 := ReplStr(SVal0, '.', ',');
          try    SVal0 := IntToStr(Round(StrToFloat(SVal0)));
          except SVal0 := '';
          end;
       end;

       {���������� ������}
       If (SVal0='') or (SVal0='0') or (Not IsIntegerStr(SVal0)) then PFile^.DeleteKey  (STab, SCol+' '+SRow)
                                                                 else PFile^.WriteString(STab, SCol+' '+SRow, SVal0);

       {������ ������ ������������}
       If Not PFile^.SectionExists(STab) then begin
          PFile^.WriteString(STab, '0', '0');
          PFile^.DeleteKey  (STab, '0');
       end;

       {������������ ���������}
       Result:=true;
    finally
    end;
end;


{==============================================================================}
{===================       ����� DAT-����� � ����       =======================}
{==============================================================================}
{===================   SPath  - ���� � �������� �����   =======================}
{===================   RESULT - ������ ����             =======================}
{==============================================================================}
function TBASE.FindMemFile(const SPath: String): PMemIniFile;
var I      : Integer;
    SPath0 : String;
begin
    Result := nil;
    SPath0 := AnsiUpperCase(SPath);
    For I:=Low(MemFile) to High(MemFile) do begin
       If AnsiUpperCase(MemFile[I].FileName) = SPath0 then begin
          Result := @MemFile[I];
          Break;
       end;
    end;
end;


{==============================================================================}
{========================  �������� MEM-�����  ================================}
{==============================================================================}
function TBASE.AddMemFile(const SPath: String): PMemIniFile;
var I : Integer;
begin
    Result := nil;
    try
       I := Length(MemFile);
       SetLength(MemFile, I + 1);
       MemFile[I] := TMemIniFile.Create(SPath);
       Result := @MemFile[I];
    finally
    end;
end;


{==============================================================================}
{========  ���������� ����� ���� � DAT-����� � ������ �� ����������  ==========}
{==============================================================================}
procedure TBASE.SaveMemFile;
var I : Integer;
begin
    For I:=Low(MemFile) to High(MemFile) do MemFile[I].UpdateFile;
end;

