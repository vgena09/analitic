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
{==================   СОХРАНЯЕТ АКТИВНУЮ ЯЧЕЙКУ ДЛЯ КНИГИ WB    ===============}
{==============================================================================}
function SaveActiveCell(const WB: Variant): Boolean;
var F    : TIniFile;
    SLog : String;
    SPos : String; // лист графа строка
begin
    {Инициализация}
    Result:=false;
    If VarIsEmpty(WB) then Exit;
    try
       {Идентификатор формы}
       SLog := WB.WorkSheets[1].Range[CELL_LOG].Value;
       If Not CmpStr(IDLOG, Copy(SLog, 1, Length(IDLOG))) then Exit;
       Delete(SLog, 1, Length(IDLOG));
       SLog:=Trim(SLog);
       If SLog='' then Exit;

       {Определяем активную ячейку}
       SPos:=IntToStr(WB.ActiveSheet.Index)+' '+
             IntToStr(WB.Application.ActiveCell.Column)+' '+
             IntToStr(WB.Application.ActiveCell.Row);

       {Запоминаем активную ячейку}
       F:=TIniFile.Create(PATH_WORK_INI);
       try     F.WriteString(INI_ACTIVE_CELL, SLog, SPos);
       finally F.Free; end;

       {Возвращаемый результат}
       Result:=true;
    finally
    end;
end;


{==============================================================================}
{==============   УСТАНАВЛИВАЕТ АКТИВНУЮ ЯЧЕЙКУ ДЛЯ КНИГИ WB    ===============}
{==============================================================================}
function LoadActiveCell(const WB: Variant): Boolean;
var F                 : TIniFile;
    S, SPos, SLog     : String;
    IPage, IRow, ICol : Integer;
    Cell              : Variant;
begin
    {Инициализация}
    Result:=false;
    If VarIsEmpty(WB) then Exit;
    try
       {Идентификатор формы}
       SLog := WB.WorkSheets[1].Range[CELL_LOG].Value;
       If Not CmpStr(IDLOG, Copy(SLog, 1, Length(IDLOG))) then Exit;
       Delete(SLog, 1, Length(IDLOG));
       SLog:=Trim(SLog);
       If SLog='' then Exit;

       {Читаем активную ячейку}
       F:=TIniFile.Create(PATH_WORK_INI);
       try     SPos:=F.ReadString(INI_ACTIVE_CELL, SLog, '');
       finally F.Free; end;
       If SPos='' then Exit;
       S := CutSlovo(SPos, 1, ' '); If Not IsIntegerStr(S) then Exit; IPage := StrToInt(S);
       S := CutSlovo(SPos, 2, ' '); If Not IsIntegerStr(S) then Exit; IRow  := StrToInt(S);
       S := CutSlovo(SPos, 3, ' '); If Not IsIntegerStr(S) then Exit; ICol  := StrToInt(S);

       {Устанавливаем активную ячейку}
       try
          WB.WorkSheets[IPage].Activate;
          Cell:=WB.WorkSheets[IPage].Cells[ICol, IRow];
          Cell.Select;
          Cell.Activate;
       except end;

       {Возвращаемый результат}
       Result:=true;
    finally
    end;
end;


end.
