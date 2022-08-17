unit MIMPORT_XLS;

interface

uses
   Winapi.Windows {Beep},
   System.Classes, System.SysUtils, System.Variants, System.IniFiles,
   Vcl.Dialogs, Vcl.Controls, IdGlobal,
   Data.DB, Data.Win.ADODB,
   MAIN, FunType;

type
   TIMPORT_XLS = class
   public
      constructor Create;
      destructor  Destroy; override;

      function  ImportXls(const SPath: String): Boolean;
      function  ImportXlsID(const WB: Variant; const IsDialog: Boolean): Boolean;
   private
      FFMAIN  : TFMAIN;
   end;

implementation

uses FunConst, FunExcel, FunIO, FunBD, FunSys, FunVcl, FunText, FunDay, FunSum,
     FunVerify, FunIni,
     MTITLE, MBASE;


{==============================================================================}
{=============================   �����������   ================================}
{==============================================================================}
constructor TIMPORT_XLS.Create;
begin
    inherited Create;

    FFMAIN  := TFMAIN(GlFindComponent('FMAIN'));
end;


{==============================================================================}
{==============================   ����������   ================================}
{==============================================================================}
destructor TIMPORT_XLS.Destroy;
begin
    // ...

    inherited Destroy;
end;


{==============================================================================}
{=====================  ����������: ������ ����� XLS  =========================}
{==============================================================================}
function TIMPORT_XLS.ImportXls(const SPath: String): Boolean;
var E: Variant;
begin
    {�������������}
    Result:=false;

    {��������� Excel}
    If Not CreateExcel then begin ErrMsg('������ �������� MS Excel!'); Exit; end;
    try

       {��������� ���� � ����������� ��������}
       If Not OpenXls(SPath, true) then begin ErrMsg('������ �������� �����: '+ExtractFileName(SPath)); Exit; end;
       try

          {������ ������}
          E      := GetExcelApplicationID;
          E.Workbooks[1].Windows[1].Visible:=true;     //��� �������� ������� ����� LOCALE_USER_DEFAULT
          Result := ImportXlsID(E.ActiveWorkbook, true);

       {��������� ����}
       finally
          E.DisplayAlerts:=false;
          If Not CloseXls then ErrMsg('������ �������� �������������� ������!');
       end;

    {��������� Excel}
    finally
       If Not CloseExcel then ErrMsg('������ �������� Excel!');
       E  := Unassigned;
    end;
end;


{==============================================================================}
{============================  ������ ����� XLS  ==============================}
{==============================================================================}
function TIMPORT_XLS.ImportXlsID(const WB: Variant; const IsDialog: Boolean): Boolean;
label Dl, Er;
var FTITLE                        : TFTITLE;
    FBASE                         : TBASE;
    LForma_                       : TLForma;
    FIni                          : TMemIniFile;
    LIni                          : TStringlist;
    FDate                         : TDate;
    SPathIni, SPathDat            : String;
    S, SLog                       : String;
    SForm, SYear, SMonth, SRegion : String;
    SPageList,  SPage             : String;
    SBlockList, SBlock            : String;
    I, IPage                      : Integer;
    IsVerify                      : Boolean;


    {==========================================================================}
    {==================   ������������ ���� �������    ========================}
    {==========================================================================}
    {==================   SBlock = 'A5:BH125'          ========================}
    {==================   SBlock = S1+':'+S2           ========================}
    {==================   SBlock = Row1,Col1,Row2,Col2 ========================}
    {==========================================================================}
    function InputBlock(const IPage: Integer; const SBlock: String): Boolean;
    var S, S1, S2              : String;
        SVal                   : Variant;
        IDTab, IDRow, IDCol    : String;
        Row1, Col1, Row2, Col2 : Integer;
        IRow, ICol             : Integer;
        AData                  : Variant;
    begin
        {�������������}
        Result:=false;
        S1 := AnsiUpperCase(CutSlovo(SBlock, 1, ':'));
        S2 := AnsiUpperCase(CutSlovo(SBlock, 2, ':'));
        If S1='' then Exit;
        If S2='' then S2:=S1;

        {����������}
        FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������: [ '+SRegion+' ] '+IntToStr(IPage)+'!'+SBlock;

        {���������� �������}
        If Not ExcelSeparatAddress(S1, Col1, Row1) then Exit;
        If Not ExcelSeparatAddress(S2, Col2, Row2) then Exit;
        If (Col1>Col2) or (Row1>Row2) then Exit;
        AData := VarArrayCreate([Row1, Row2, Col1, Col2], varVariant);
        AData := WB.WorkSheets[IPage].Range[S1, S2].Value;      // ������������ �-���: ��� S1=S2 ���������� �� ������ � ��������, � �.�. � Unassigned; �������� ������/����� �������: AData[IRow, ICol] - �� �����

        try
           {������������� ��� ������ �������}
           For IRow := Row1 to Row2 do begin
              For ICol := Col1 to Col2 do begin
                 {��������� ����� ������ � ����� A1}
                 S:=ExcelCombineAddress(ICol, IRow);
                 If S='' then Exit;
                 {������ ��� ��� ������� ������: 5/4/21/0/0}
                 S:=FIni.ReadString(INI_PAGE+SPage, S, '');
                 If S='' then Continue;
                 {��������� ��� �� ������������}
                 IDTab := CutSlovo(S, 1, CH_SPR);
                 IDCol := CutSlovo(S, 2, CH_SPR);
                 IDRow := CutSlovo(S, 3, CH_SPR);
                 {���������� �������� � DAT-���� �� �������}
                 If S1<>S2 then begin
                    SVal := AData[IRow-Row1+1, ICol-Col1+1];       // AData[IRow, ICol] - �� �����
                 end else begin
                    If Not VarIsEmpty(AData) then SVal := AData else SVal := '';
                 end;
                 If VarType(SVal)=varError then SVal:='';
                 If Not FBASE.SetVal(SPathDat, IDTab, IDCol, IDRow, SVal) then Exit;
              end;
           end;
           Result:=true;
        finally
           {����������� ������}
           VarClear(AData);
        end;
    end;


    {==========================================================================}
    {========       ��������� ������-������ �� ������������        ============}
    {==========================================================================}
    {========  SPeriod - '(�� 9 ������� 2008 �. � ��������� ...'   ============}
    {========  SPeriod - '�� 2 ������ 2008 �.' '�� 2 ������ 2008�.'============}
    {========  SPeriod - '�� ��������� 2008 �.'                    ============}
    {========  SPeriod - '�� 2008 ���'                             ============}
    {==========================================================================}
    function SeparatOldPeriod(const SPeriod: String; var SYear, SMonth: String): Boolean;
    var S, S1 : String;
        I     : Word;
    begin
        {�������������}
        Result := false;
        SYear  := '';
        SMonth := '';
        S      := Trim(SPeriod);

        {S = '(�� 9 ������� 2008' ��� '�� 2008'}
        I := Pos(' �.' , S); If I>4 then Delete(S, I, Length(S));
        I := Pos(' ���', S); If I>4 then Delete(S, I, Length(S));
        I := Pos('�.' , S);  If I>3 then Delete(S, I, Length(S));
        I := Pos('���', S);  If I>3 then If Pos('���������', S)=0 then Delete(S, I, Length(S));
        S := Trim(S);

        {�������� ���}
        S1 := TokCharEnd(S, ' ');
        If (Not IsIntegerStr(S1)) or (Length(S1)<>4) then Exit;
        SYear := S1;

        {�������� �����}
        {S = '(�� 9 �������' ��� '��'}
        TokChar(S, ' ');
        If S<>'' then begin
           I := MonthOldStrToInd(S);   If (I<Low(MonthList)) or (I>High(MonthList)) then Exit;
           SMonth := MonthList[I];
        end else begin
           SMonth := MonthList[High(MonthList)];
        end;

        {������������ ���������}
        Result:=true;
    end;

begin
    {�������������}
    Result := false;

    {�������� ��� �����}
    SLog := WB.WorkSheets[1].Range[CELL_LOG].Value;

    {**************************************************************************}
    {******************   ������������ ����������   ***************************}
    {**************************************************************************}
    If CmpStr(IDLOG,  Copy(SLog, 1, Length(IDLOG))) then begin
       {������������ �������������}
       Delete(SLog, 1, Length(IDLOG));
       SLog:=Trim(SLog);

       {���������� ������� ���, �����, ������}
       SYear   := WB.WorkSheets[1].Range[CELL_YEAR].Value;
       SMonth  := WB.WorkSheets[1].Range[CELL_MONTH].Value;
       SRegion := WB.WorkSheets[1].Range[CELL_REGION].Value;
       Goto Dl;
    end;

    {**************************************************************************}
    {**************************   ���������� ���  *****************************}
    {**************************************************************************}
    // If WB.WorkSheets.Count<5 then Goto Er;  701 ����� �������� 4 �����
    SLog := '';
    try
       {�������� ��� ����� � ����������������}
       If FFMAIN.FindForms(@LForma_, false, true, -1, -1, '', '', '', '', true, 0) then begin
          For I:=High(LForma_) downto Low(LForma_) do begin
             {���������� �������������}
             SForm := ReadWBCells(WB, LForma_[I].FCellID);
             If Not CmpStr(SForm, LForma_[I].FStrID) then Continue;

             {������ ������ � ������}
             S       := ReadWBCells(WB, LForma_[I].FCellPeriod);
             SRegion := ReadWBCells(WB, LForma_[I].FCellRegion);

             {���� ������ ����� � ����������� �� �������}
             SeparatOldPeriod(S, SYear, SMonth);
             FDate := SDatePeriod(SYear, SMonth);
             If LForma_[I].FBegin > 0 then If LForma_[I].FBegin > FDate then Continue;
             If LForma_[I].FEnd   > 0 then If LForma_[I].FEnd   < FDate then Continue;

             {������ ����� ������}
             SLog := IntToStr(LForma_[I].FCounter);
             Break;
          end;
       end;
    finally
       SetLength(LForma_, 0);
    end;
    If SLog <> '' then Goto Dl;

    {����� �� ���������������}
Er: ErrMsg('���� �� ������� ��� �������� ������!');
    Exit;

Dl: {������������ �����}
    If (Pos(' ', SLog)>0) or (Not IsIntegerStr(SLog)) then begin ErrMsg('��� ����� �� ���������: '+SLog); Exit; end;

    {���������� ini-���� ����� - SPathIni}
    SForm    := FFMAIN.FormsCounterToCaption(StrToInt(SLog));
    If SForm='' then begin ErrMsg('��� ����� �� ���������: '+SLog); Exit; end;
    S        := FFMAIN.FormsCounterToFile(StrToInt(SLog));
    SPathIni := ChangeFileExt(S, '.ini');

    {��������� ini-����}
    If Not FileExists(SPathIni) then begin ErrMsg('�� ������ ini-����: '+SPathIni); Exit; end;
    FIni  := TMemIniFile.Create(SPathIni);
    FBASE := TBASE.Create;
    LIni  := TStringList.Create;
    try
       {������ ������ ��������������� ������}
       If Not FFMAIN.GetListTabIni(@LIni, SPathIni) then begin ErrMsg('������ ������ ������ ������ �� INI-�����!'+CH_NEW+SPathIni); Exit; end;

       {������������-�������������}
       FTITLE:=TFTITLE.Create(nil);
       try     If Not FTITLE.Execute(SYear, SMonth, SRegion, SForm, false, IsDialog) then Exit;
       finally FTITLE.Free;
       end;
       FFMAIN.Repaint;

       {�������������� � ������������� ������ ������ ������}
       SPathDat:=FFMAIN.PathRegion(SYear, SMonth, SRegion);
       //YMR := FFMAIN.SetYMR(SYear, SMonth, SRegion, '');
       //If FFMAIN.IsTableDatIncludeIni(SPathIni, SPathDat) then begin
       //   If MessageDlg('���� ������ �������� �������, ������� ��� ������� ������ ����� �������� �� �����! ����������?', mtWarning, [mbYes, mbNo], 0)<>mrYes
       //   then Exit;
       //end;

       {������� �������� dat-������}
       S := FFMAIN.CreateRegions(SYear, SMonth, SRegion);
       If S='' then Exit;

       {������������ ���������}
       WriteLocalString(INI_SELECT, INI_SELECT_PATH, S);

       {������ ������� �������� XLS-�������}
       IsVerify := ReadLocalBool(INI_SET, INI_SET_IMPORT_VERIFY_XLS, true);

       {������ ������ ������}
       SPageList:=FIni.ReadString(INI_PARAM, INI_PARAM_PAGES, '');

       {������������� ��������� �����}
       SPage:=CutBlock(SPageList, ',');
       While SPage<>'' do begin
          {������������ ������ �����}
          If Not IsIntegerStr(SPage) then begin ErrMsg('�������� ����� ����� ['+SPage+'] !'); Exit; end;
          IPage := StrToInt(SPage);

          {����������� �������� ���� ��� ���������� ��������}
          FFMAIN.Repaint;

          {������ ������� �����}
          SBlockList:=FIni.ReadString(INI_PAGE+SPage, INI_PAGE_IN, '');

          {�������: ���������� ������}
          // WB.WorkSheets[IPage].Range['A2','A2'].Insert(3);

          {������������� ��� ������� ����� �������� �����}
          SBlock:=CutBlock(SBlocklist, ',');
          While SBlock<>'' do begin
             If Not InputBlock(IPage, SBlock) then begin ErrMsg('������ ������ ����� ['+SPage+'!'+SBlock+'] !'); Exit; end;
             SBlock:=CutBlock(SBlockList, ',');
          end;

          {��������� ����}
          SPage:=CutBlock(SPageList, ',');
       end;

       {��������� �������}
       FBASE.SaveMemFile;
       FBASE.ClearCash;

       {������������� ������� � ���������}
       For I:=0 to LIni.Count-1 do SumTable(LIni[I], SYear, SMonth, SRegion);

       {��������� �������}
       If IsVerify then VerifyTable(@LIni, SYear, SMonth, SRegion, true);
    finally
       LIni.Free;
       FBASE.Free;
       FIni.Free;
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
    end;

    {������������ ���������}
    Result:=true;
end;

end.
