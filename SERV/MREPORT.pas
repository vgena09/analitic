unit MREPORT;

interface

uses
   Winapi.Windows,
   System.Classes, System.SysUtils, System.Variants, System.IniFiles,
   System.DateUtils, System.Math,
   Vcl.Dialogs, Vcl.Controls, IdGlobal, Vcl.ComCtrls, Vcl.OleServer,
   Data.DB, Data.Win.ADODB, Excel2000,
   MAIN, MPARAM, MBASE, FunType;

type
   TAnaliz = record
      PLParam               : PListParam;    // ��������� �� ������ ���������� ������������ ���������, ��������������� ��� �������
   // LParam                : TListParam;    // ������ ���������� ������������ ���������, ��������������� ��� �������
      LRegions              : TStringList;   // ������ ������������� ��������     ��� ������� �������: ������ ��������� �������� - ���������: 1 - �����, 2 - �������, 3 - ����������
      IStep                 : Integer;       // ��� ������� (� �������)    [-12]
      ILoop                 : Integer;       // ����� ��������             [2]
      IRowBefore, IRowAfter : Integer;       // ����� ����� � ����� � ������� (������ ��� ����������� ����� ����������)
      FirstColDat           : Integer;       // ������ ������� � �������      (������ ��� ����������� ����� ����������)
      IPeriod               : Integer;       // ������� ������     ILoop * IStep
   end;
   PAnaliz = ^TAnaliz;

   TREPORT = class
   public
      constructor Create;
      destructor  Destroy; override;

      function  CreateReport(const SYear_, SMonth_, SRegion_, SFormula_ : String;
                             const IDReport_, IDMatrix_, CorrMonth_     : Integer;
                             const IsNew : Boolean): Boolean;
   private
      FFMAIN                           : TFMAIN;
      FPARAM                           : TFPARAM;
      FBASE                            : TBASE;
      WB                               : Variant;
      FIni                             : TMemIniFile;
      Analiz                           : TAnaliz;

      SYear, SMonth, SRegion, SFormula : String;
      IDReport, IDMatrix, CorrMonth    : Integer;

      SType, SPage, SBlock             : String;
      IPage                            : Integer;

      IsCtrl                           : Boolean;             // ������ �� Crtl �� ����� ������� �������

      function  CreatePage(const SPage_ : String): Boolean;
      function  CreateBlockStat: Boolean;
      function  CreateBlockAnal: Boolean;

      function  ValueCell(const SKey: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
      function  ValueCellStep(const SBlock_: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
      procedure SetInfo(const SDopText: String);
   end;

implementation

uses FunConst, FunExcel, FunIO, FunBD, FunSys, FunVcl, FunInfo, FunText, FunDay, FunSum;

{$INCLUDE MREPORT_STAT}
{$INCLUDE MREPORT_ANAL}


{==============================================================================}
{=============================   �����������   ================================}
{==============================================================================}
constructor TREPORT.Create;
begin
    inherited Create;

    {�������������}
    FFMAIN          := TFMAIN(GlFindComponent('FMAIN'));
    FPARAM          := TFPARAM(FFMAIN.FFPARAM);
    FBASE           := TBASE.Create;
    Analiz.LRegions := TStringList.Create;
    FFMAIN.Repaint;
end;


{==============================================================================}
{==============================   ����������   ================================}
{==============================================================================}
destructor TREPORT.Destroy;
begin
    {���������������}
    FBASE.Free;
    Analiz.LRegions.Clear;
    Analiz.LRegions.Free;
    WB := Unassigned;
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';

    inherited Destroy;
end;


{==============================================================================}
{=========                    ������� ����� EXCEL                      ========}
{==============================================================================}
{=========  CorrMonth - ��������� ������ ��� �������� ������ ������    ========}
{=========  �������   - ������������ ��������                          ========}
{==============================================================================}
function TREPORT.CreateReport(const SYear_, SMonth_, SRegion_, SFormula_ : String;
                              const IDReport_, IDMatrix_, CorrMonth_     : Integer;
                              const IsNew : Boolean): Boolean;
var LPag                                    : TStringList;
    IPage, I                                : Integer;
    SPathIni, SPathSrc, SPathDect, SPathDat : String;
begin
    {�������������}
    Result    := false;
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������������ ������: �������������';
    SYear     := SYear_;
    SMonth    := SMonth_;
    SRegion   := SRegion_;
    SFormula  := SFormula_;
    IDReport  := IDReport_;
    IDMatrix  := IDMatrix_;
    CorrMonth := CorrMonth_;
    IsCtrl    := GetKeyState(VK_LCONTROL);
    If (Not IsIntegerStr(SYear)) or (MonthStrToInd(SMonth)<=0) then begin ErrMsg('������������ ������: ['+SYear+' '+SMonth+']!'); Exit; end;
    If SRegion = ''                                            then begin ErrMsg('�� ��������� ������!'); Exit; end;
    If (IDMatrix < 0) or (IDMatrix > 2)                        then begin ErrMsg('������������ �������� IDMatrix!'); Exit; end;
    Analiz.PLParam := @FPARAM.LParam;

    {����-��������}
    SPathSrc := FFMAIN.FormsCounterToFile(IDReport);
    If Not FileExists(SPathSrc) then begin ErrMsg('�� ������ ������ ������!'+CH_NEW+SPathSrc); Exit; end;

    {���� ��������}
    SPathIni := ChangeFileExt(SPathSrc, '.ini');
    If Not FileExists(SPathIni) then begin ErrMsg('�� ������ ���� ��������!'+CH_NEW+SPathIni); Exit; end;

    {����-����������}
    SPathDect := CopyFileToTemp(SPathSrc);
    If SPathDect='' then begin ErrMsg('�� ��������� ���� ������!'+CH_NEW+SPathDect); Exit; end;

    {�������������� � ������������� ������ ������ ������}
    If IsNew then begin
       SPathDat := FFMAIN.PathRegion(SYear, SMonth, SRegion);
       If FFMAIN.IsTableDatIncludeIni(SPathIni, SPathDat) then begin
          If MessageDlg('���� ������ �������� �������, ������� ��� ���������� ������������ ������ ����� �������� �� �����! ����������?', mtWarning, [mbYes, mbNo], 0)<>mrYes
          then Exit;
       end;
    end else CorrMonth := 0;

    {������������� INI}
    FIni := TMemIniFile.Create(SPathIni);
    LPag := TStringList.Create;
    try
       {������ ������}
       If Not SeparatMStr(FIni.ReadString(INI_PARAM, INI_PARAM_PAGES, ''), @LPag, ', ') then begin ErrMsg('������ � ������ ������ ������!'); Exit; end;

       {�������������� � ������� ������}
       If (Not IsNew) and (LPag.Count >= 50) then begin
          If MessageDlg('����� �������� '+ IntToStr(LPag.Count) +' �������������� ������� � ��� ��� �������� ����� ������������ �����. ����������?', mtWarning, [mbYes, mbNo], 0)<>mrYes
          then Exit;
       end;

       try
          {����������}
          FFMAIN.PBar.Position := 0;
          FFMAIN.PBar.Visible  := true;
          FFMAIN.Repaint;

          {������������� EXCEL}
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' ������������ ������: �������� EXCEL';
          FFMAIN.PBar.Position := 2;
          FFMAIN.WB.Disconnect;
          FFMAIN.XLS.Disconnect;
          FFMAIN.XLS.ConnectKind          := ckNewInstance;                     // ckRunningOrNew; - ����������� �.�. ��������
          FFMAIN.XLS.Connect;
          FFMAIN.XLS.AutoQuit             := false;
          FFMAIN.XLS.EnableEvents         := true;
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' ������������ ������: �������� ���������';
          FFMAIN.PBar.Position := 5;
          FFMAIN.XLS.Workbooks.Add(SPathDect, LOCALE_USER_DEFAULT);
          If FFMAIN.XLS.Workbooks.Count = 0 then begin ErrMsg('������ �������� ������!'); Exit; end;
          If (IDMatrix = 0) and (GetMSOfficeVer(msExcel) > 11) then FFMAIN.RunMacroSafe('��������.Workbook_Open'); // ��� Office > 2003
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' ������������ ������: ��������� ���������';
          FFMAIN.PBar.Position := 8;
          FFMAIN.XLS.Visible[LOCALE_USER_DEFAULT]             := false;         // ��� Office > 2003
          Variant(FFMAIN.XLS.Application).UseSystemSeparators := false;         // �� �������� ��� Office > 2003 (������� �� ������ ������)
          Variant(FFMAIN.XLS.Application).DecimalSeparator    := ',';           // ...
          Variant(FFMAIN.XLS.Application).ThousandsSeparator  := '�';           // ...
          FFMAIN.XLS.UseSystemSeparators := false;                              // ���������� ������ �� > 2003
          FFMAIN.XLS.DecimalSeparator    := ',';                                // ...
          FFMAIN.XLS.ThousandsSeparator  := '�';                                // ...

          FFMAIN.WB.ConnectTo(FFMAIN.XLS.ActiveWorkbook);
          WB := FFMAIN.XLS.Workbooks[1];
          WB.Application.EnableEvents    := false;
          WB.Application.CellDragAndDrop := false;                              // ��������� �������������� �����

          {���������� � ����� ������� ���, �����, ������}
          WB.WorkSheets[1].Unprotect;
          WB.WorkSheets[1].Range[CELL_YEAR].Value    := SYear;
          WB.WorkSheets[1].Range[CELL_MONTH].Value   := SMonth;
          WB.WorkSheets[1].Range[CELL_REGION].Value  := SRegion;
          WB.WorkSheets[1].Range[CELL_CAPTION].Value := FFMAIN.FormulaToCaption(SFormula, false);

          {��������� ��������� ����� ������}
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' ������������ ������: ���������� ���������';
          FFMAIN.PBar.Position := 10;
          For IPage:=0 to LPag.Count-1 do begin
             I := StrToInt(LPag[IPage]);
             If IsNew and (CorrMonth = 0) then begin
                {������� �����}
                try If (IDMatrix = 0) and FIni.ReadBool(INI_PAGE+LPag[IPage], INI_PAGE_PROTECT, true) then begin
                       WB.WorkSheets[I].EnableSelection := xlUnlockedCells;
                       WB.WorkSheets[I].Protect('');
                    end;
                except end;
             end else begin
                {��������� �����}
                If LPag.Count > 1 then FFMAIN.PBar.Position := 10 + (85 * IPage Div (LPag.Count-1))
                                  else FFMAIN.PBar.Position := 55;
                If Not CreatePage(LPag[IPage]) then begin ErrMsg('������ ��������� ����� ������ � '+LPag[IPage]+'!'); Exit; end;
             end;
             {��� Office > 2003 �������������}
             If GetMSOfficeVer(msExcel) > 11 then begin
                WB.WorkSheets[I].EnableFormatConditionsCalculation := true;     // ��������� ��������� ��������� ��������������
                WB.WorkSheets[I].EnableCalculation := true;                     // ��������� ��������
             end;
          end;

          {����������}
          FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������������ ������: ���������� ��������';
          FFMAIN.PBar.Position := 100;

          {������� ������}
          If IDMatrix>0 then WB.Application.ActiveWindow.DisplayWorkbookTabs := true
                        else WB.Application.ActiveWindow.DisplayWorkbookTabs := FIni.ReadBool(INI_PARAM, INI_PARAM_PAGELINKS, false);

          {��������� ��������� ������� � EXCEL}
          WB.Application.EnableEvents:=true;

          {������������ ������� ���������, ���������}
          If (IDMatrix = 0) then FFMAIN.RunMacroSafe('RunAfterCreate');

          {��� ������� ��������� ��� ���� ��������}
          If (IDMatrix > 0) then begin
             FFMAIN.XLS.DisplayFullScreen[LOCALE_USER_DEFAULT]:=false;
             FFMAIN.XLS.WindowState[0] := xlMaximized;
          end;

          {�������������� ��������� ����� � ������}
          LoadActiveCell(FFMAIN.XLS.Workbooks[1]);

          {��������� EXCEL}
          FFMAIN.XLS.Visible[LOCALE_USER_DEFAULT] := true;
          If GetMSOfficeVer(msExcel)>=11 then SetForegroundWindow(Variant(FFMAIN.XLS.DefaultInterface).Hwnd);
          FFMAIN.XLS.Workbooks[1].Saved[LOCALE_USER_DEFAULT]:=true;
          FFMAIN.XLS.UserControl := true;

          {������������ ���������}
          Result:=true;
       finally
          If Result then begin
             Result:=IS_ADMIN and (IDMatrix=0) and (Not FFMAIN.IsFormReadOnly(IDReport_));
          end else begin
             try FFMAIN.WB.Close(false); except end;
             FFMAIN.WB.Disconnect;
             FFMAIN.XLS.Disconnect;
          end;

          FFMAIN.PBar.Visible  := false;
          FFMAIN.PBar.Position := 0;
       end;

   finally
       LPag.Free;
       FIni.Free;
       WB     := Unassigned;
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
   end;
end;


{==============================================================================}
{=======================  ������� �������� ������ EXCEL  ======================}
{==============================================================================}
function TREPORT.CreatePage(const SPage_ : String): Boolean;
var SBlockList : String;
begin
    {�������������}
    Result := false;
    SPage  := SPage_;
    If Not IsIntegerStr(SPage) then begin ErrMsg('������������ ����� �����: '+SPage+'!'); Exit; end;
    IPage  := StrToInt(SPage);
    FFMAIN.Repaint;

    {������������ ������ �����}
    If IPage > Integer(WB.WorkSheets.Count) then begin ErrMsg('������������ ����� �����: '+SPage+'!'); Exit; end;

    {��� ������}
    SType := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_TYPE, '');

    {������ �������������� ������ �������� �����}
    Case IDMatrix of
    0, 2 : SBlockList := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_OUT, '');
    1    : SBlockList := FIni.ReadString(INI_PAGE+SPage, INI_PAGE_IN, '');
    else begin ErrMsg('������������ �������� IDMatrix!'); Exit; end;
    end;

    {������������ ����� �������� �����}
    WB.WorkSheets[IPage].Unprotect;
    try
       SBlock := CutBlock(SBlocklist, ',');
       While SBlock<>'' do begin
          If SType = '' then begin If Not CreateBlockStat then Exit; end
                        else begin If Not CreateBlockAnal then Exit; end;
          SBlock := CutBlock(SBlockList, ',');
       end;
    finally
       If (IDMatrix = 0) and FIni.ReadBool(INI_PAGE+SPage, INI_PAGE_PROTECT, true) then begin
          WB.WorkSheets[IPage].EnableSelection := xlUnlockedCells;
          WB.WorkSheets[IPage].Protect('');
       end;
    end;

    {������������ ���������}
    Result:=true;
end;



{******************************************************************************}
{**************************  ��������������� �������  *************************}
{******************************************************************************}


{==============================================================================}
{==========================  ���������� �������� ������   =====================}
{==============================================================================}
function TREPORT.ValueCell(const SKey: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
var SKey0  : String;
    SBlock : String;
    SVal   : String;
begin
    {�������������}
    Result := 0;
    SKey0  := SKey;

    //{��� ������������� ������������ ������ � ������}
    //If CorrPeriod <> 0  then SKey0 := SKey0 + '/' + IntToStr(CorrPeriod);
    //If CorrRegion <> '' then SKey0 := SKey0 + '/' + CorrRegion;

    try
       {��� ������� ���� ������ ����� - �������� ������ ���������}
       SBlock := CutModulChar(SKey0, '[', ']');
       If SBlock<>'' then begin
          While SBlock<>'' do begin
             SVal   := FloatToStr(ValueCellStep(SBlock, CorrPeriod, CorrRegion));
             SKey0  := ReplModulChar(SKey0, SVal, '[', ']');
             SBlock := CutModulChar(SKey0,        '[', ']');
          end;

          {��������� ��������}
          Result:=CalcStr(SKey0);

       {��� ���������� ������ - ���������� ��������}
       end else begin
          Result := ValueCellStep(SKey, CorrPeriod, CorrRegion);
       end;
    finally
    end;
end;


{==============================================================================}
{==========================  ���������� �������� ������   =====================}
{==============================================================================}
function TREPORT.ValueCellStep(const SBlock_: String; const CorrPeriod: Integer; const CorrRegion: String): Extended;
var IDTab, IDRow, IDCol, IDReg, IDPer : String;
    SBlock                            : String;
    SPath                             : String;
begin
    {�������������}
    Result := 0;
    SBlock := SBlock_;
    try
       IDTab  := CutBlock(SBlock, CH_SPR);
       If CmpStr(IDTab, '���') then begin
          SBlock := SFormula+CH_SPR+SBlock;
          IDTab  := CutBlock(SBlock, CH_SPR);
       end;
       IDCol  := CutBlock(SBlock, CH_SPR);
       IDRow  := CutBlock(SBlock, CH_SPR);
       IDPer  := CutBlock(SBlock, CH_SPR); If IDPer='' then IDPer:='0';
       IDReg  := CutBlock(SBlock, CH_SPR);
       SBlock := IDTab+CH_SPR+IDCol+CH_SPR+IDRow;
       If CorrPeriod <> 0 then SBlock := SBlock+CH_SPR+IntToStr(StrToInt(IDPer)+CorrPeriod)
                          else SBlock := SBlock+CH_SPR+IDPer;
       If CorrRegion <>'' then SBlock := SBlock+CH_SPR+CorrRegion
                          else SBlock := SBlock+CH_SPR+IDReg;
       // ������ �� �����������: � ����� ����������� ��������� ������� ���������� ������
       SPath  := FFMAIN.PathDat(SYear, SMonth, SRegion, SBlock);
       Result := FBASE.GetVal(SPath, IDTab, IDCol, IDRow);
    finally end;
end;


{==============================================================================}
{================================  ����������  ================================}
{==============================================================================}
procedure TREPORT.SetInfo(const SDopText: String);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������������ ������: [ '+SRegion+' ] '+SPage+'!'+SBlock+' '+SDopText;
    FFMAIN.StatusBar.Repaint;
end;


end.
