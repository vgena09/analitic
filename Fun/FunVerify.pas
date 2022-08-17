unit FunVerify;

interface
uses
   System.Classes, System.SysUtils, System.Variants, System.IniFiles,
   Vcl.Dialogs, Vcl.Controls,
   Data.DB, Data.Win.ADODB,
   IdGlobal,
   FunType;

function VerifyTable(const PLTab: PStringList; const SYear, SMonth, SRegion: String; const IsAllVisible: Boolean): Boolean;

implementation

uses FunConst, FunSys, FunVcl, FunText, FunDay, FunIO, FunBD,
     MAIN, MINFO, MBASE;


{==============================================================================}
{=========================  �������� ������ ������  ===========================}
{==============================================================================}
function VerifyTable(const PLTab: PStringList; const SYear, SMonth, SRegion: String; const IsAllVisible: Boolean): Boolean;
label Nx;
var T, T_TAB, T_COL, T_ROW : TADOTable;
    DS_TAB                 : TDataSource;
    FFMAIN                 : TFMAIN;
    FFINFO                 : TFINFO;
    FBASE                  : TBASE;
    LExcept                : TStringList;
    YMR, YMR1, YMR3        : String;
    STab, SCol, SRow       : String;
    I, ITab                : Integer;
    IMonth                 : Integer;
    IsError, IsVal2        : Boolean;
    IVal2                  : Extended;

    {��������� ������� YMR � YMRPrev}
    procedure VerifyStep(const YMRPrev: String; const STab, SCol, SRow: String);
    var IVal1 : Extended;
    begin
        {������ ���������� ��������}
        IVal1 := FBASE.GetValYMR(YMRPrev, STab, SCol, SRow);
        If IVal1=0 then Exit;

        {������ ������� ��������}
        If Not IsVal2 then begin
           IVal2  := FBASE.GetValYMR(YMR, STab, SCol, SRow);
           IsVal2 := true;
        end;

        {���������� ���������}
        If IVal1>IVal2 then begin
           IsError := true;
           FFINFO.AddInfo('������: '+
                  CutLongStr(FFMAIN.RowsNumericToCaption(SRow, ITab), 80)+' \ '+
                  CutLongStr(FFMAIN.ColsNumericToCaption(SCol, ITab), 80)+':  '+
                  FloatToStr(IVal1)+' > '+FloatToStr(IVal2), ICO_ERR);
        end;
    end;
begin
    {�������������}
    Result  := false;
    If PLTab=nil then Exit;
    If Not IsIntegerStr(SYear) then Exit;
    IMonth  := MonthStrToInd(SMonth);
    IsError := false;
    FFMAIN  := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' �������� ������ ...';
    FFINFO  := TFINFO.Create(nil);
    LExcept := TStringList.Create;
    T_TAB   := TADOTable.Create(nil);
    T_COL   := TADOTable.Create(nil);
    T_ROW   := TADOTable.Create(nil);
    DS_TAB  := TDataSource.Create(nil);
    FBASE   := TBASE.Create;
    try
       {������������� ������ ����������}
       T:=TADOTable.Create(nil);
       try
          If Not OpenBD(@FFMAIN.LBank[0].BD, '', '', [@T], [TABLE_EXCEPT]) then begin ErrMsg('������ �������� ������� ����������!'); Exit; end;
          try
             T.First;
             While Not T.Eof do begin
                LExcept.Add(T.FieldByName(EXCEPT_ADRES).AsString);
                T.Next;
             end;
          finally T.Close;
          end;
       finally T.Free;
       end;

       {������������� ������}
       If Not OpenBD(@FFMAIN.LBank[0].BD, '', '', [@T_TAB,       @T_COL,     @T_ROW],
                                                  [TABLE_TABLES, TABLE_COLS, TABLE_ROWS]) then begin ErrMsg('������ �������� ������!'); Exit; end;
       DS_TAB.DataSet := T_TAB;
       SetDBConnect(@T_COL, @DS_TAB, F_COUNTER, TABLE_TABLES);
       SetDBConnect(@T_ROW, @DS_TAB, F_COUNTER, TABLE_TABLES);

       {������������� ����������}
       FFINFO.Caption              := '�������� ������';
       FFINFO.IsStop               := false;
       FFINFO.AReport.Enabled      := false;
       FFINFO.BtnClose1.ModalResult := mrNone;
       FFINFO.AClose.Caption       := '��������';
       FFINFO.Show;
       FFINFO.Repaint;

       {������������� �������}
       For I:=0 to PLTab^.Count-1 do begin
          {�������������}
          STab := PLTab^[I];
          If Not IsIntegerStr(STab) then Continue;
          ITab := StrToInt(Stab);
          SetDBFilter(@T_TAB, '['+F_COUNTER+']='+STab);
          If T_TAB.RecordCount<>1 then Continue;
          YMR  := FFMAIN.SetYMR(SYear, SMonth, SRegion, STab);
          YMR1 := FFMAIN.SetYMR(SYear, SMonth, SRegion, STab+CH_SPR+' '+CH_SPR+' '+CH_SPR+'-1');
          YMR3 := FFMAIN.SetYMR(SYear, SMonth, SRegion, STab+CH_SPR+' '+CH_SPR+' '+CH_SPR+'-3');
          FFINFO.AddInfo('�������: '+FFMAIN.TablesCounterToCaption(STab), ICO_INFO);

          {������������� ������� � ������}
          T_COL.First;
          While Not T_COL.Eof do begin
             SCol:=T_COL.FieldByName(F_NUMERIC).AsString;
             T_ROW.First;
             While Not T_ROW.Eof do begin
                SRow   := T_ROW.FieldByName(F_NUMERIC).AsString;
                IsVal2 := false;

                {�������� �� ����������}
                If LExcept.IndexOf(STab+' '+SCol+' '+SRow)>=0 then Goto Nx;

                {�������� ������ 1}
                Case IMonth of
                2..12: VerifyStep(YMR1, STab, SCol, SRow);
                end;

                {�������� ������ 2}
                Case IMonth of
                6, 9, 12: VerifyStep(YMR3, STab, SCol, SRow);
                end;

            Nx: {��������� ������}
                T_ROW.Next;
             end;
             If FFINFO.IsStop then Break;
             T_COL.Next;
          end;
          If FFINFO.IsStop then Break;
       end;

       {��������� ����������}
       FFINFO.BtnClose1.ModalResult := mrOk;
       FFINFO.AClose.Caption       := '�������';

       {��������� ��������}
       If FFINFO.IsStop then begin
          If IsError then FFINFO.AddInfo('�������� ��������: ������� ������', ICO_WARN)
                     else FFINFO.AddInfo('�������� ��������', ICO_WARN);
       end else begin
          If IsError then FFINFO.AddInfo('�������� ���������: ������� ������', ICO_WARN)
                     else FFINFO.AddInfo('�������� ���������: ������ �� ����������', ICO_OK);
       end;

       {��������� ����� ���� ��� ������ ��� ��� ������������ ��������}
       FFINFO.Hide;
       FFINFO.AReport.Enabled := true;
       If IsAllVisible or IsError then FFINFO.ShowModal;

       Result:=true;
    finally
       CloseBD(nil, [@T_TAB, @T_COL, @T_ROW]);
       DS_TAB.Free;
       T_TAB.Free;
       T_COL.Free;
       T_ROW.Free;
       LExcept.Free;
       FFINFO.Free;
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
       FBASE.Free;
    end;
end;

end.

