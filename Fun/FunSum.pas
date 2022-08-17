unit FunSum;

interface
uses
   System.Classes, System.SysUtils, System.IniFiles, System.Variants,
   Vcl.Dialogs, Vcl.Controls,
   Data.DB, Data.Win.ADODB,
   IdGlobal,
   FunType;

function  SumTable(const STab, SYear, SMonth, SRegion: String): Boolean;

implementation

uses FunConst, MAIN,
     FunBD, FunText, FunDay, FunIO, FunSys, FunIni;

{==============================================================================}
{===================     ������������ ������ �������    =======================}
{==============================================================================}
{============   SRegion - �������� ������� � ���������� ��������  =============}
{==============================================================================}
function SumTable(const STab, SYear, SMonth, SRegion: String): Boolean;
var FFMAIN        : TFMAIN;
    LRegion_      : TLRegion;
    Sr            : TSearchRec;
    DatPeriod     : TDate;
    IYear, IMonth : Word;
    IDRegion      : Integer;

    SListParent   : TStringList;
    IDParent      : Integer;
    SParentRegion : String;
    SPathDat      : String;
    IndParent     : Integer;

    ExistsSection : Boolean;
    S, S1, S2     : String;
    I             : Integer;

    function SumStep(const SRegionStep: String): Boolean;
    var FDat        : TMemIniFile;
        SList       : TStringList;
        SDat        : String;
        S, S0, SKey : String;
        I           : Integer;
        IVal        : Extended;
    begin
        {�������������}
        Result:=false;

        {�������������� dat-���� �������� �������}
        SDat:=FFMAIN.PathDat(SYear, SMonth, SRegionStep, STab);
        If Not FileExists(SDat) then Exit;
        {������ ������ ������������}
        If Not FFMAIN.IsExistsTab(SDat, STab) then Exit;

        FDat  := TMemIniFile.Create(SDat);
        SList := TStringList.Create;
        try
           {���������� ���������� �����}
           FDat.ReadSectionValues(STab, SList);

           {���������� ������������}
           For I:=0 to SList.Count-1 do begin
              S    := SList[I];                        //  S = '1 2=5'
              SKey := TokChar(S, '=');                 //  S = '5'
              {���� �������� ���������}
              If IsIntegerStr(S) then begin
                 {���� ��������������� ������������ ������}
                 IndParent:=SListParent.IndexOfName(SKey);
                 {���� ������������ ������ ����������}
                 If IndParent>=0 then begin
                    S0:=SListParent.Values[SKey];
                    If IsIntegerStr(S0) then IVal:=StrToFloat(S0) else IVal:=0;
                    SListParent.Values[SKey]:=FloatToStr(StrToFloat(S)+IVal);
                 {���� ������������ ������ �����������}
                 end else begin
                    SListParent.Add(SKey+'='+S);
                 end;
              end;
           end;

           {������������ ���������}
           Result:=true;
        finally
           SList.Free;
           FDat.Free;
        end;
    end;

begin
    {�������������}
    Result := false;
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));

    If Not IsIntegerStr(SYear) then Exit;
    IYear  :=StrToInt(SYear);
    IMonth := MonthStrToInd(SMonth);
    If IMonth<=0 then Exit;
    DatPeriod := IDatePeriod(IYear, IMonth);
    If DatPeriod=0 then Exit;

    S2 := CutModulStr (SRegion,     SUB_RGN1, SUB_RGN2);
    S1 := ReplModulStr(SRegion, '', SUB_RGN1, SUB_RGN2);

    SListParent := TStringList.Create;
    try
        {������ ��������� ����}
        If Not FFMAIN.FindRegions(@LRegion_, true, -1, S1, -1, -1, DatPeriod) then Exit;
        IDRegion := LRegion_[0].FCounter;


        {**********************************************************************}
        {*****   ������������ ��������� ��������� ��������   ******************}
        {**********************************************************************}
        If S2<>'' then begin
           {�������� � ������ ������������� ����}
           SParentRegion := S1;
           IDParent      := IDRegion;

           {����������}
           FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������ �������: '+STab+'! [ '+SParentRegion+' ]';

           {���� � ��������� dat-�����}
           SPathDat := FFMAIN.PathDat(SYear, SMonth, S1, STab);

           {���� ������ ��������� ������}
           S:=ExtractFilePath(SPathDat);
           try
              If FindFirst(S+S1+SUB_RGN1+'*', faDirectory, Sr) = 0 then begin
                 {������� ���������� �������� ������ dat-�����}
                 FFMAIN.DelTab(SPathDat, STab);

                 {���������� ������}
                 repeat S:=ExtractFileNameWithoutExt(Sr.Name);
                        SumStep(S);
                 until  FindNext(Sr) <> 0;

                 {��������� ����������� ������ dat-����� ������������� ����}
                 If Not FFMAIN.SetTabVal(SPathDat, STab, @SListParent, true) then Exit;
              end;
           finally FindClose(Sr); end;


        {**********************************************************************}
        {*****   ������������ ��������   **************************************}
        {**********************************************************************}
        end else begin
           {������������� ������� ���������� ��� ������}
           If FFMAIN.IsTableNoStandart(STab) then begin
              Result:=true;
              Exit;
           end;

           {������ ������������� ����}
           IDParent:=FFMAIN.RegionsCounterToParent(IDRegion);
           If IDParent=-1 then Exit;

           {���� ������ ����� ����, �� ������ �����������}
           If IDParent=0 then begin
              Result:=true;
              Exit;
           end;

           {�������� ������������� ����}
           SParentRegion:=FFMAIN.RegionsCounterToCaption(IDParent);
           If SParentRegion='' then Exit;

           {����������}
           FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������ �������: '+STab+'! [ '+SParentRegion+' ]';

           {������� ���������� �������� ������ dat-�����}
           SPathDat:=FFMAIN.PathDat(SYear, SMonth, SParentRegion, STab);
           FFMAIN.DelTab(SPathDat, STab);

           {�������� ������� � ���������� ������������ �����}
           If Not FFMAIN.FindRegions(@LRegion_, false, -1, '', IDParent, -1, DatPeriod) then Exit;

           {���������� ��������� ��������� �������}
           ExistsSection := true;
           For I:=Low(LRegion_) to High(LRegion_) do begin
              {���� ������� ������� �����������, �� �������� ���� ������� ������}
              ExistsSection := SumStep(LRegion_[I].FCaption);
              {���� �������������� ������������, �� ������ ����������}
              If FFMAIN.IsTableNoStandart(STab) then ExistsSection:=true;
              If Not ExistsSection then Break;
           end;

           {��������� ����������� ������ dat-����� ������������� ����}
           If ExistsSection then begin
              If Not FFMAIN.SetTabVal(SPathDat, STab, @SListParent, true) then Exit;
           {������� ������ dat-����� ������������� ����}
           end else begin
              If Not FFMAIN.DelTab(SPathDat, STab) then Exit;
           end;
        end;
        {**********************************************************************}
     finally
        SListParent.Free;
        SetLength(LRegion_, 0);
     end;

     {����������� ����� ������� ���������� ������}
     If IDParent <> IMAIN_REGION then Result := SumTable(STab, SYear, SMonth, SParentRegion)
                                 else Result := true;
end;

end.

