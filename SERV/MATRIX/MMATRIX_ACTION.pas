{==============================================================================}
{=========================   ����������� ACTION   =============================}
{==============================================================================}
procedure TFMATRIX.EnablAction;
var BRun, BOpen, BClear : Boolean;
    STab, SCol, SRow    : String;
begin
    {�������������}
    BRun   := true;
    BOpen  := true;
    BClear := true;
    try
       {������� ������ ���� ��������}
       If EBlockKod.Text = '' then BRun := false;

       {������ ���� ���� �����}
       If EBlockIn.Text = '' then BRun := false;
       
       {�� ����� ���� ������ ������ ����� �����}
       If Pos(',', EBlockIn.Text) > 0 then BRun := false;
       If Pos(' ', EBlockIn.Text) > 0 then BRun := false;

       {������ ���� ���� ������}
       If EBlockOut.Text = '' then BRun := false;

       {��� ������ ���� c����������}
       If Not SeparatKod(EBlockKod.Text, STab, SRow, SCol) then BRun:=false;
       If (STab='') or (SRow='') or (SCol='') then BRun:=false;

    finally
       ARun.Enabled   := BRun;
       AOpen.Enabled  := BOpen;
       AClear.Enabled := BClear;
    end;
end;


{==============================================================================}
{===========================   ACTION: ���������   ============================}
{==============================================================================}
procedure TFMATRIX.ARunExecute(Sender: TObject);
var F                          : TIniFile;
    SList                      : TStringList;
    STab, SRow, SCol, SPage, S : String;
    FirstRow, FirstCol         : Integer;

    {==========================================================================}
    {=====   ���������� �����   ===============================================}
    {==========================================================================}
    function IndexBlock(const SBlock: String): Boolean;
    var SBlock1, SBlock2 : String;
        SKey, SVal       : String;
        Row1, Col1       : Integer;
        Row2, Col2       : Integer;
        IRow, ICol       : Integer;
    begin
        {�������������}
        Result:=false;

        {������ � ����� �������}
        SBlock1 := AnsiUpperCase(CutSlovo(SBlock, 1, ':'));
        SBlock2 := AnsiUpperCase(CutSlovo(SBlock, 2, ':'));
        If SBlock1='' then Exit;
        If SBlock2='' then SBlock2:=SBlock1;

        {���������� ������� � ������� A1}
        If Not ExcelSeparatAddress(SBlock1, Col1, Row1) then Exit;
        If Not ExcelSeparatAddress(SBlock2, Col2, Row2) then Exit;
        If (Col1>Col2) or (Row1>Row2) then Exit;

        {������������� ��� ������ �������}
        For ICol := Col1 to Col2 do begin
           For IRow := Row1 to Row2 do begin
              {��������� ����� ������ � ����� A1}
              SKey := ExcelCombineAddress(ICol, IRow);
              If SKey='' then Exit;
              {��������� ���}
              SVal := STab+CH_SPR+IntToStr(ICol-Col1+FirstCol)+CH_SPR+IntToStr(IRow-Row1+FirstRow);
              {���������� ���}
              F.WriteString(INI_PAGE+SPage, SKey, SVal);
           end;
        end;

        {������������ ���������}
        Result:=true;
    end;

begin
    {�������������}
    If Not SeparatKod(EBlockKod.Text, STab, SRow, SCol) then Exit;
    FirstRow := StrToInt(SRow);
    FirstCol := StrToInt(SCol);
    SPage    := EPage.Text;

    F     := TIniFile.Create(SPathIni);
    SList := TStringList.Create;
    try
       {���������� ������ �������}
       S:=F.ReadString(INI_PARAM, INI_PARAM_TABLES, '');
       SeparatMStr(S, @SList, ', ');
       If SList.IndexOf(STab)<0 then begin
          If S<>'' then S:=S+', ';
          S:=S+STab;
          F.WriteString(INI_PARAM, INI_PARAM_TABLES, S);
       end;

       {���������� ������ �����}
       S:=F.ReadString(INI_PARAM, INI_PARAM_PAGES, '');
       SeparatMStr(S, @SList, ', ');
       If SList.IndexOf(SPage)<0 then begin
          If S<>'' then S:=S+', ';
          S:=S+SPage;
          F.WriteString(INI_PARAM, INI_PARAM_PAGES, S);
       end;

        {���������� ������� ����� � ��}
        If EBlockIn.Text<>'' then begin
           S:=F.ReadString(INI_PAGE+SPage, INI_PAGE_IN, '');
           SeparatMStr(S, @SList, ', ');
           If SList.IndexOf(EBlockIn.Text)<0 then begin
              If S<>'' then S:=S+', ';
              S:=S+EBlockIn.Text;
              F.WriteString(INI_PAGE+SPage, INI_PAGE_IN, S);
           end;
        end;

        {���������� ������� ������ �� ��}
        If EBlockOut.Text<>'' then begin
           S:=F.ReadString(INI_PAGE+SPage, INI_PAGE_OUT, '');
           SeparatMStr(S, @SList, ', ');
           If SList.IndexOf(EBlockOut.Text)<0 then begin
              If S<>'' then S:=S+', ';
              S:=S+EBlockOut.Text;
              F.WriteString(INI_PAGE+SPage, INI_PAGE_OUT, S);
           end;
       end;
       
       {���������� ������� �����}
       If Not IndexBlock(EBlockIn.Text) then Exit;

       {����������}
       ShowMessage('������� ���������������: '+CH_NEW+ExtractFileName(SPathIni));

       {������� �����}
       EBlockIn.Text  := '';
       EBlockOut.Text := '';
       EnablAction;

       {��������� ������ �����}
       If EBlockIn.Enabled then EBlockIn.SetFocus;
    finally
       If SList<>nil then SList.Free;
       F.Free;
    end;
end;


{==============================================================================}
{============================   ACTION: ��������   ============================}
{==============================================================================}
procedure TFMATRIX.AClearExecute(Sender: TObject);
var F        : TIniFile;
    Section  : String;
begin
    {�������������}
    Section:=INI_PAGE+EPage.Text;
    F:=TIniFile.Create(SPathIni);
    try
       If Not F.SectionExists(Section) then begin
          MessageDlg('������� ��� �������� � '+EPage.Text+' �� ����������!', mtError, [mbOk], 0);
          Exit;
       end;

       If MessageDlg('����������� ������� �������� ��� ����� � '+EPage.Text+'.',
       mtWarning, [mbYes, mbNo], 0)<>mrYes then Exit;

       F.EraseSection(Section);
       MessageDlg('������� ��� �������� � '+EPage.Text+' �������!', mtInformation, [mbOk], 0);
    finally
       F.Free;
    end;
end;


{==============================================================================}
{============================   ACTION: �������   =============================}
{==============================================================================}
procedure TFMATRIX.AOpenExecute(Sender: TObject);
begin
    StartAssociatedExe(SPathINI, SW_SHOWNORMAL);
end;


{==============================================================================}
{=====================  �������� ���� �������� �����    =======================}
{==============================================================================}
procedure TFMATRIX.EBlockOutButtonClick(Sender: TObject);
begin
    EBlockOut.Text := '';
    SOut           := '';
    EChange(Sender);
end;


{==============================================================================}
{===================  ���������� � ���� �������� �����    =====================}
{==============================================================================}
procedure TFMATRIX.EBlockInAltButtonClick(Sender: TObject);
begin
    EBlockOut.Text := EBlockIn.Text;
    SOut           := '';
    EChange(Sender);
end;


{==============================================================================}
{==========  ���������������� ��������� ����� ����� � ������ ������  ==========}
{==============================================================================}
procedure TFMATRIX.EBlockInButtonClick(Sender: TObject);
begin
    EChange(Sender);
end;


