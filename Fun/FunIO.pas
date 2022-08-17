unit FunIO;

interface
uses
   System.SysUtils, System.Variants, System.Classes, System.IniFiles,
   Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
   Data.DB, Data.Win.ADODB,
   FunType;

   function  SaveActiveCell(const WB: Variant): Boolean;
   function  LoadActiveCell(const WB: Variant): Boolean;

implementation

uses FunConst, FunText, FunBD, FunExcel, FunSys, FunDay, FunIni;


{==============================================================================}
{==================   ��������� �������� ������ ��� ����� WB    ===============}
{==============================================================================}
function SaveActiveCell(const WB: Variant): Boolean;
var F    : TIniFile;
    SLog : String;
    SPos : String; // ���� ����� ������
begin
    {�������������}
    Result:=false;
    If VarIsEmpty(WB) then Exit;
    try
       {������������� �����}
       SLog := WB.WorkSheets[1].Range[CELL_LOG].Value;
       If Not CmpStr(IDLOG, Copy(SLog, 1, Length(IDLOG))) then Exit;
       Delete(SLog, 1, Length(IDLOG));
       SLog:=Trim(SLog);
       If SLog='' then Exit;

       {���������� �������� ������}
       SPos:=IntToStr(WB.ActiveSheet.Index)+' '+
             IntToStr(WB.Application.ActiveCell.Column)+' '+
             IntToStr(WB.Application.ActiveCell.Row);

       {���������� �������� ������}
       F:=TIniFile.Create(PATH_WORK_INI);
       try     F.WriteString(INI_ACTIVE_CELL, SLog, SPos);
       finally F.Free; end;

       {������������ ���������}
       Result:=true;
    finally
    end;
end;


{==============================================================================}
{==============   ������������� �������� ������ ��� ����� WB    ===============}
{==============================================================================}
function LoadActiveCell(const WB: Variant): Boolean;
var F                 : TIniFile;
    S, SPos, SLog     : String;
    IPage, IRow, ICol : Integer;
    Cell              : Variant;
begin
    {�������������}
    Result:=false;
    If VarIsEmpty(WB) then Exit;
    try
       {������������� �����}
       SLog := WB.WorkSheets[1].Range[CELL_LOG].Value;
       If Not CmpStr(IDLOG, Copy(SLog, 1, Length(IDLOG))) then Exit;
       Delete(SLog, 1, Length(IDLOG));
       SLog:=Trim(SLog);
       If SLog='' then Exit;

       {������ �������� ������}
       F:=TIniFile.Create(PATH_WORK_INI);
       try     SPos:=F.ReadString(INI_ACTIVE_CELL, SLog, '');
       finally F.Free; end;
       If SPos='' then Exit;
       S := CutSlovo(SPos, 1, ' '); If Not IsIntegerStr(S) then Exit; IPage := StrToInt(S);
       S := CutSlovo(SPos, 2, ' '); If Not IsIntegerStr(S) then Exit; IRow  := StrToInt(S);
       S := CutSlovo(SPos, 3, ' '); If Not IsIntegerStr(S) then Exit; ICol  := StrToInt(S);

       {������������� �������� ������}
       try
          WB.WorkSheets[IPage].Activate;
          Cell:=WB.WorkSheets[IPage].Cells[ICol, IRow];
          Cell.Select;
          Cell.Activate;
       except end;

       {������������ ���������}
       Result:=true;
    finally
    end;
end;


end.
