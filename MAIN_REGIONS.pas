{******************************************************************************}
{************************  ��������������� �������  ***************************}
{************************   �� ������ � ���������   ***************************}
{******************************************************************************}


{==============================================================================}
{=====================   ������������ ��������� �������   =====================}
{==============================================================================}
{===========             ������ ������� �� �����������                  =======}
{==============================================================================}
{===========  ���� ������ �� ����� ����������� �����������:             =======}
{===========  LTab - ������ ������������ � ��������/����������� ������  =======}
{===========  LTab = nil - ��� ��������� ������� �������                =======}
{==============================================================================}
function TFMAIN.IsModifyRegion(const SYear, SMonth, SRegion: String; const PLTab: PStringList): Boolean;
var LRegion_           : TLRegion;
    SRegion1, SRegion2 : String;
    FDate              : TDate;
    I, Ind             : Integer;

    LTabSub            : TStringList;
    LFind              : array of String;
begin
    {�������������}
    Result   := false;
    If (SYear='') or (SMonth='') or (PLTab=nil) then Exit;
    SRegion1 := ReplModulStr(SRegion, '', SUB_RGN1, SUB_RGN2);
    SRegion2 := CutModulStr (SRegion,     SUB_RGN1, SUB_RGN2);
    FDate    := SDatePeriod(SYear, SMonth);

    try
       {������� � ������ ��� -  ����������� ���������}
       If Not FindRegions(@LRegion_, true, -1, SRegion1, -1, -1, FDate) then Exit;
       Ind := LRegion_[0].FCounter;

       {���� ���� ���� ����������� ������� - �������� ������������}
       Result := (PLTab^.Count > 0);
       For I:=0 to PLTab^.Count-1 do begin
          If Not IsTableNoStandart(PLTab^[I]) then begin
             Result:=false;
             Break;
          end;
       end;
       If Result then Exit;

       {������ ����� ����������� ���������� - ����������� ���������}
       If FindRegions(@LRegion_, true, -1, '', Ind, -1, FDate) then Exit;
    finally
       SetLength(LRegion_, 0);
    end;

    {��������� ������ - ����������� ���������}
    If SRegion2<>'' then begin
       Result:=true;
       Exit;

    {������ �� ����� ����������� �����������}
    end else begin
       try
          {������ ����� ��������� ���������� - ������������� ����� ������}
          If FindDataYMR(@LFind, SYear+'\'+SMonth+'\'+SRegion1+SUB_RGN1+'*', true) > 0 then begin
             {�������������}
             If PLTab^.Count=0 then Exit;
             LTabSub := TStringList.Create;
             try
                {������������� ��� ��������� ����������}
                For I:=Low(LFind) to High(LFind) do begin
                   {�������� ������ ������ ���������� ����������}
                   If Not GetListTabDat(@LTabSub, LFind[I]) then Exit;
                   {��������� ��������� �� ������ ��������� �� ����� �� �������� ������}
                   For Ind:=0 to PLTab^.Count-1 do begin
                      If LTabSub.IndexOf(PLTab^[Ind])>=0 then Exit;
                   end;
                end;
                Result:=true;
             finally
                LTabSub.Free;
             end;

          {������ �� ����� ��������� ����������� - ����������� ���������}
          end else begin
             Result:=true;
          end;
       finally
          SetLength(LFind, 0);
       end;
    end;
end;


{==============================================================================}
{======================  ������ �������� ��� ���������  =======================}
{==============================================================================}
function TFMAIN.ListRegionsAnalitic(const PList: PStringList; const SBaseRegion: String;
                                    const IsApparat, IsMainRegion: Boolean;
                                    const SYear, SMonth: String): Boolean;
var LRegion_ : TLRegion;
    FDate    : TDate;
    I, Ind   : Integer;
begin
    {�������������}
    Result := false;
    If PList=nil then Exit;
    FDate  := SDatePeriod(SYear, SMonth);
    PList^.Clear;
    try
       {������ �������� �������}
       If Not FindRegions(@LRegion_, true, -1, SBaseRegion, -1, -1, FDate) then begin ErrMsg('������ ����������� ������� ������� ['+SBaseRegion+'] !'); Exit; end;
       Ind := LRegion_[0].FCounter;

       {������ ��������� ��������}
       If FindRegions(@LRegion_, false, -1, '', Ind, -1, FDate) then begin
          For I:=1 to Length(LRegion_)-1 do PList^.Add(LRegion_[I].FCaption);
       end;

       {��������� �������}
       If IsApparat and (Length(LRegion_)>0) then PList^.Add(LRegion_[0].FCaption);

       {��������� ������� ������}
       If IsMainRegion then PList^.Add(SBaseRegion);

       {������������ ���������}
       Result:=true;

    finally
       SetLength(LRegion_, 0);
    end;
end;


{==============================================================================}
{===========  ������������ ���������� ������ ��������� ��������   =============}
{==============================================================================}
{===========    ����������� [...], �����������, ���������, ��     =============}
{==============================================================================}
{===========         ���� SYEAR, SMONTH = '' - ��� �������        =============}
{==============================================================================}
function TFMAIN.ListRegionsRangeUp(const PList: PStringList; const SBaseRegion: String;
                                   const SYear, SMonth: String): Boolean;
var LRegion_ : TLRegion;
    FDate    : TDate;
    S1, S2   : String;
    IParent  : Integer;
begin
    {�������������}
    Result := false;
    If PList=nil then Exit;
    PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);
    S2    := CutModulStr(SBaseRegion, SUB_RGN1, SUB_RGN2);
    S1    := ReplModulStr(SBaseRegion, '', SUB_RGN1, SUB_RGN2);
    If S1='' then Exit;

    try
       {���� ������� � ������ ���}
       If Not FindRegions(@LRegion_, true, -1, S1, -1, -1, FDate) then Exit;
       IParent := LRegion_[0].FParent;

       {C��������}
       If S2<>'' then PList^.Add(SBaseRegion);

       {������� ������}
       PList^.Add(S1);

       {������������ �������}
       While IParent>0 do begin
          If Not FindRegions(@LRegion_, true, IParent, '', -1, -1, FDate) then Exit;
          PList^.Add(LRegion_[0].FCaption);
          IParent := LRegion_[0].FParent;
       end;
    finally
       SetLength(LRegion_, 0);
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{===========   ������������ ���������� ������ ��������� ��������   ============}
{==============================================================================}
{===========         ���������, �����������, ��������, ...         ============}
{==============================================================================}
{===========         ���� SYEAR, SMONTH = '' - ��� �������         ============}
{==============================================================================}
function TFMAIN.ListRegionsRangeDown(const PList: PStringList; const SBaseRegion: String;
                                     const SYear, SMonth: String): Boolean;
var LRegion_ : TLRegion;
    FDate    : TDate;
    S1, S2   : String;

    procedure AddReg(const IParent: Integer);
    var LRegion_ : TLRegion;
        I        : Integer;
    begin
        try
           {�������� ������ ��������� ��������}
           If Not FindRegions(@LRegion_, false, -1, '', IParent, -1, FDate) then Exit;
           {������������� ������}
           For I:=0 to Length(LRegion_)-1 do begin
              PList^.Add(LRegion_[I].FCaption);
              AddReg(LRegion_[I].FCounter);
           end;
        finally
           SetLength(LRegion_, 0);
        end;
    end;

begin
    {�������������}
    Result := false;
    If PList=nil then Exit;
    PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);
    S2    := CutModulStr(SBaseRegion,      SUB_RGN1, SUB_RGN2);
    S1    := ReplModulStr(SBaseRegion, '', SUB_RGN1, SUB_RGN2);
    If (S1='') or (S2<>'') then Exit;

    try
       {���� ������� � ������ ���}
       If Not FindRegions(@LRegion_, true, -1, S1, -1, -1, FDate) then Exit;

       {�������� ������� � �������� �������}
       PList^.Add(LRegion_[0].FCaption);
       AddReg(LRegion_[0].FCounter);
    finally
       SetLength(LRegion_, 0);
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{======================  ������ �������� ������ ������  =======================}
{==============================================================================}
function TFMAIN.ListRegionsLevel(const PList: PStringList; const SBaseRegion: String;
                                 const SYear, SMonth: String): Boolean;
var LRegion_  : TLRegion;
    FDate     : TDate;
    I, ILevel : Integer;
begin
    {�������������}
    Result := false;
    If PList=nil then Exit;                PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);   If FDate=0 then Exit;
    If IsSubRegion(SBaseRegion) then Exit;

    try
       {������� �������� �������}
       If Not FindRegions(@LRegion_, true, -1, SBaseRegion, -1, -1, FDate) then Exit;
       ILevel := LRegion_[0].FGroup;

       {�������� ������� � ����� �������}
       If Not FindRegions(@LRegion_, false, -1, '', -1, ILevel, FDate) then Exit;

       {��������� ������ ��������}
       For I:=Low(LRegion_) to High(LRegion_) do PList^.Add(LRegion_[I].FCaption);
    finally
       SetLength(LRegion_, 0);
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{==========================   ����� ������ ��������   =========================}
{==============================================================================}
function TFMAIN.CopyListRegions(const PList: PStringList; const SYear, SMonth: String): Boolean;
var FDate : TDate;
    I     : Integer;
begin
    {�������������}
    Result := false;
    If PList=nil then Exit;
    PList^.Clear;
    FDate := SDatePeriod(SYear, SMonth);
    If FDate=0 then Exit;

    {��������� ������}
    For I:=Low(LRegion) to High(LRegion) do begin
       If LRegion[I].FBegin > 0 then If LRegion[I].FBegin > FDate then Continue;
       If LRegion[I].FEnd   > 0 then If LRegion[I].FEnd   < FDate then Continue;
       PList^.Add(LRegion[I].FCaption);
    end;

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{=========================    ������� DAT-�����    ============================}
{=========================    ���� ��� TREEPATH    ============================}
{==============================================================================}
function TFMAIN.CreateRegions(const SYear, SMonth, SRegion: String): String;
var SList   : TStringlist;
    SPath   : String;
    SFile   : String;
    Result0 : String;
    I       : Integer;
begin
    {�������������}
    Result  := '';
    Result0 := '';
    SPath   := PATH_BD_DATA+SYear+'\'+SMonth+'\';
    SList   := TStringList.Create;
    try
       {�������� ������������� ���������� ������ ��������}
       If Not ListRegionsRangeUp(@SList, SRegion, SYear, SMonth) then Exit;
       If SList.Count=0 then Exit;

       {������������� �������}
       For I:=0 to SList.Count-1 do begin
          Result0 := SList[I]+'\'+Result0;
          SFile   := SPath+SList[I]+'.dat';
          If Not CreateEmptyFile(SFile) then begin ErrMsg('������ �������� ����� ['+SFile+']!'); Exit; end;
       end;

    finally
       SList.Free;
    end;

    {������������ ���������}
    Result:=SYear+'\'+SMonth+'\'+Result0;
end;


{==============================================================================}
{========   ��������� ������������ �������� ������� ��� ����������    =========}
{==============================================================================}
function TFMAIN.VerifyRegionName(const SRegionName: String): Boolean;
var S : String;
    I : Integer;
begin
    {�������������}
    Result := false;
    {�������� �������� �������}
    S := ReplModulStr(SRegionName, '', SUB_RGN1, SUB_RGN2);

    {��������� �������� �������}
    For I:=Low(LRegion) to High(LRegion) do begin
       If LRegion[I].FCaption = S then begin
          Result:=true;
          Break;
       end;
    end;
end;


{==============================================================================}
{================   �������� �� �������� ������� �����������    ===============}
{==============================================================================}
function TFMAIN.IsSubRegion(const SRegion: String): Boolean;
begin
    Result := (Pos(SUB_RGN1, SRegion) > 0);
end;


(*
{==============================================================================}
{===========================   ������� �������   ==============================}
{==============================================================================}
function TFMAIN.RegionLevel(const SRegion: String; const FDate: TDate): Integer;
var LRegion_ : TLRegion;
    IParent  : Integer;
    IResult  : Integer;
begin
    {�������������}
    Result  := -1;
    IResult := -1;
    If FDate=0 then Exit;
    If IsSubRegion(SRegion) then Exit;
    SetLength(LRegion_, 0);
    try
       {������ �������� �������� �������}
       If Not FindRegions(@LRegion_, true, -1, SRegion, -1, -1, FDate) then begin ErrMsg('������ ����������� ������� ������������� ������� ['+SRegion+'] !'); Exit; end;
       IParent  := LRegion_[0].FParent;
       IResult  := 0;
       {������������� �������}
       While IParent > 0 do begin
          If Not FindRegions(@LRegion_, true, IParent, '', -1, -1, FDate) then begin ErrMsg('������ ����������� ������� ������������� ������� ['+SRegion+'] !'); Exit; end;
          IParent  := LRegion_[0].FParent;
          IResult  := IResult + 1;
       end;
       Result := IResult;
    finally
       SetLength(LRegion_, 0);
    end;
end;
*)

