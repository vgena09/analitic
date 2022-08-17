{==============================================================================}
{==========================   ПОДКЛЮЧЕНИЕ К EXCEL   ===========================}
{==============================================================================}
function TFMATRIX.ConnectExcel(const SPath: String): Boolean;
begin
    {Инициализация}
    Result:=false;

    {Закрытие предыдущего отчета}
    DisconnectExcel(true);

    try
       {Открытие EXCEL}
       XLS.ConnectKind := ckNewInstance;
       XLS.Connect;
       XLS.AutoQuit:=false;
       XLS.Workbooks.Add(SPath, LOCALE_USER_DEFAULT);
       If XLS.Workbooks.Count = 0 then Exit;
       WB.ConnectTo(XLS.ActiveWorkbook);

       {Видимость и доступность отчета}
       XLS.Visible[LOCALE_USER_DEFAULT] := true;
       XLS.UserControl := true;
       XLS.Workbooks[1].Saved[LOCALE_USER_DEFAULT]:=true;

       {Возвращаемый результат}
       Result:=true;
    finally
    end;
end;


{==============================================================================}
{=======================   ПРИ ИЗМЕНЕНИИ SELECTED   ===========================}
{==============================================================================}
procedure TFMATRIX.WBSheetSelectionChange(ASender: TObject; const Sh: IDispatch; const Target: ExcelRange);
var WB                     : Variant;
    Col1, Col2, Row1, Row2 : Integer;
    S, SRegion             : String;
    V                      : Variant;
begin
    {Инициализация}
    If (Not IsIn) and (Not IsOut) then Exit;
    V    := XLS.Selection[LOCALE_USER_DEFAULT];
    WB   := XLS.Workbooks[1];
    Col1 := V.Column;
    Row1 := V.Row;
    If (Row1=0) or (Col1=0) then Exit;
    Row2 := Row1 + V.Rows.Count - 1;
    Col2 := Col1 + V.Columns.Count - 1;

    {Определяем выделенную область SRegion}
    SRegion := ExcelCombineAddress(Col1, Row1);
    If SRegion='' then Exit;
    S := ExcelCombineAddress(Col2, Row2);
    If (S<>'') and (S<>SRegion) then SRegion:=SRegion+':'+S;

    {Записываем индекс листа и выделенную область}
    EPage.Value :=WB.ActiveSheet.Index;
    If IsIn  then begin
       EBlockIn.Text := SRegion;
       If EBlockIn.ButtonDown then begin
          EBlockOut.Text := EBlockIn.Text;
          SOut           := '';
       end;
    end;
    If IsOut then begin
       If SOut='' then EBlockOut.Text := SRegion
                  else EBlockOut.Text := SOut+', '+SRegion;
    end;

    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{========================   ПЕРЕД ЗАКРЫТИЕМ EXCEL   ===========================}
{==============================================================================}
procedure TFMATRIX.WBBeforeClose(ASender: TObject; var Cancel: WordBool);
begin
    {Отключение от EXCEL}
    DisconnectExcel(false);
end;


{==============================================================================}
{==========================   ОТКЛЮЧЕНИЕ ОТ EXCEL   ===========================}
{==============================================================================}
procedure TFMATRIX.DisconnectExcel(const IsClose: Boolean);
begin
    try
       If XLS.Workbooks.Count=0 then Exit;
       XLS.DisplayAlerts[0]:=false;           // 1 - ошибка в 2010
       XLS.Workbooks[1].Saved[LOCALE_USER_DEFAULT]:=true;
       WB.Disconnect;
       If IsClose then XLS.Quit;
       XLS.Disconnect;
    finally
    end;
end;

