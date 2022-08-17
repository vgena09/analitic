{******************************************************************************}
{**************************  РАБОТА С TREE_FORM  ******************************}
{******************************************************************************}


{==============================================================================}
{========================   ОБНОВЛЯЕТ  TREEFORM    ============================}
{==============================================================================}
procedure TFMAIN.RefreshTreeForm;
begin
    Repaint;
    TreeForm.Items.BeginUpdate;
    try
       {Очистка формы}
       TreeForm.OnChange:=nil;
       TreeForm.Items.Clear;
       TreeForm.OnChange:=TreeFormChange;

       {Обновление формы}
       TreeFormChange(nil, nil);
       
    finally
       TreeForm.Items.EndUpdate;
    end;

    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{==========================  СОБЫТИЕ: ON_CHANGE  ==============================}
{==============================================================================}
{========================== Sender=nil - очистка ==============================}
{==============================================================================}
procedure TFMAIN.TreeFormChange(Sender: TObject; Node: TTreeNode);
var FBASE                       : TBASE;
    N                           : TTreeNode;
    LFind                       : array of String;
    IFind                       : Integer;
    SList, SList2               : TStringList;
    LForma_                     : TLForma;
    I, NData, NIco, ITab, IForm : Integer;
    SCaption, SForm, SIni, YMR  : String;
    S, SCol, SRow, SExt         : String;

    function VerifyForm: Boolean;
    var NPath : TTreeNode;
    begin
        {Инициализация}
        Result:=false;

        {Проверяем наличие файла формы и ini-файла}
        If Not FileExists(SForm) then Exit;
        If Not FileExists(SIni)  then Exit;

        {Информация о выделении в TreePath}
        If Not SetSelPath then Exit;

        {Выделяем базовые год, месяц и регион формы}
        NPath:=TreePath.Selected;
        If NPath=nil then Exit;
        If NPath.Level<2 then Exit;

        {Проверяем наличие хотя бы одной секции-таблицы}
        Result := IsTableDatIncludeIniYMR(SIni, SelPath.SYear+'\'+SelPath.SMonth+'\'+SelPath.SRegion+'.dat');
    end;

    function GetCaption(const PQ: PADOTable; const SFormula: String): String;
    var S: String;
    begin
        {Инициализация}
        Result:=SFormula;
        try
           S:=CutModulChar(Result, '[', ']');
           While S<>'' do begin
              S:=PQ^.FieldByName(S).AsString;
              Result:=ReplModulChar(Result, S, '[', ']');
              S:=CutModulChar(Result, '[', ']');
           end;
        finally
           Result:=CutLongStr(Result, 100);
        end;
    end;

begin
    {Инициализация}
    StatusBar.Panels[STATUS_KOD].Text := '';

    {Сохраняем выделение}
    If Sender<>nil then SaveTreeSelect(@TreeForm, INI_SELECT, INI_SELECT_FORM);

    {Проверка TreePath: если нет dat-файла, то незачем строить структуру}
    If Not SetSelPath then Exit;
    YMR := SetYMR(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '');
    If FindDataYMR(nil, YMR, false) = 0 then Exit;

    {Инициализация интерфейса}
    TreeForm.OnChange:=nil;
    TreeForm.Items.BeginUpdate;
    try
       {Информация о выделении в TreeForm}
       If Node <> nil then begin
          StatusBar.Panels[STATUS_MAIN].Text := ' ' + Node.Text;
          If Node.Count > 0 then Exit;    // Узлы, имеющие зависимые узлы, не обрабатываются
          NData := Integer(Node.Data);
          NIco  := Node.ImageIndex;
          If NData < 0 then Exit;         // Узлы с Data < 0 не обрабатываются
       end else begin
          NData := 0;
          NIco  := 0;
       end;

       {***********************************************************************}
       {****  Выделена папка или ничего не выделено: формируем список форм ****}
       {***********************************************************************}
       If NIco = 0 then begin
          try
             {Ищем формы, зависимые от NForm}
             If Not FindForms(@LForma_, false, false, -1, NData, SelPath.SRegion, '', '', SelPath.SMonth, false, SelPath.Date) then begin
                {Если нет форм: бледная иконка и выход}
                If Node<>nil then If Node.ImageIndex=0 then SetNodeState(@Node, TVIS_CUT);
                Exit;
             end;

             {Просматриваем отобранные формы}
             For IForm:=Low(LForma_) to High(LForma_) do begin
                {Пути к файлам формы}
                SForm := PathFileForm(LForma_[IForm].SBankPref, LForma_[IForm].FFile);
                SIni  := ChangeFileExt(SForm, '.ini');

                {Определяем заголовок узла}
                SCaption := CutLongStr(LForma_[IForm].FCaption, 140);
                If SCaption='' then Continue;

                {Проверяем допустимость текущей формы и иконку}
                I := LForma_[IForm].FIcon;
                If (Not VerifyForm) and (I<>0) then Continue;

                {Формируем узел}
                N:=TreeForm.Items.AddChildObject(Node, SCaption, Pointer(LForma_[IForm].FCounter));

                {Выбираем иконку}
                If I>0 then begin
                   I := ICO_FORM_MIN + ((I - 1) * 2);
                   If I > ICO_FORM_MAX then I := ICO_FORM_MAX;
                end;
                N.ImageIndex    := I;
                N.SelectedIndex := N.ImageIndex+1;
             end;
          finally
             {Освобождаем память}
             SetLength(LForma_, 0);
          end;


          {********************************************************************}
          {***********************  Произвольные файлы  ***********************}
          {********************************************************************}
          If (Sender=nil) and (Node=nil) and (SelPath.SRegion<>'') then begin
             try
                If FindDataYMR(@LFind, SelPath.SPath+'*', true) > 0 then begin
                   For IFind:=Low(LFind) to High(LFind) do begin
                      S    := ExtractFileNameWithoutExt(LFind[IFind]);
                      SExt := AnsiUpperCase(ExtractFileExt(LFind[IFind]));
                      {Выбираем иконку}
                      I := -1;
                      If CmpStr(SExt, '.XLS') then I := ICO_FORM_FREE_XLS;
                      If CmpStr(SExt, '.PDF') then I := ICO_FORM_FREE_PDF;
                      If CmpStr(SExt, '.MDI') then I := ICO_FORM_FREE_MDI;
                      If I > -1 then begin
                         N := TreeForm.Items.AddChildObject(nil, S, Pointer(BankPrefixToIndex(PathToBankPrefix(LFind[IFind]))));
                         N.ImageIndex    := I;
                         N.SelectedIndex := N.ImageIndex+1;
                      end;
                   end;
                end;
             finally SetLength(LFind, 0);
             end;
          end;

       end;


       {***********************************************************************}
       {*************   Выделен отчет: формируем список таблиц   **************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_REPORT  then begin
          {Инициализация}
          SForm := FormsCounterToFile(NData);
          SIni  := ChangeFileExt(SForm, '.ini');
          SList := TStringList.Create;
          try
             {Читаем из ini-файла список допустимых таблиц}
             If Not GetListTabIni(@SList , SIni) then Exit;
             {Просматриваем список допустимых таблиц}
             For I:=0 to SList.Count-1 do begin
                {Ищем хоть одну таблицу в dat-файле}
                If IsExistsRowYMR(YMR, SList[I], '') then begin
                   {Заголовок отчета}
                   SCaption := TablesCounterToCaption(SList[I]);
                   If SCaption='' then Continue;
                   {Добавляем узел}
                   N := TreeForm.Items.AddChildObject(Node, SCaption, Pointer(StrToInt(SList[I])));
                   N.ImageIndex    := ICO_FORM_DETAIL_TABLE;
                   N.SelectedIndex := ICO_FORM_DETAIL_TABLE+1;
                end;
             end;
          finally
             SList.Free;
          end;
       end;


       {***********************************************************************}
       {************   Выделена таблица: формируем список строк   *************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_TABLE  then begin
          {Инициализация}
          SList  := TStringList.Create;
          SList2 := TStringList.Create;
          try
             {Читаем из dat-файла ключи выбранной таблицы (секции)}
             If Not GetTabKeyYMR(@SList, YMR, IntToStr(NData)) then Exit;
             For I:=0 to SList.Count-1 do begin
                SList[I]:=Format('%3s', [CutSlovo(SList[I], 2, ' ')]);
             end;
             SList.Sort;
             {Просматриваем список ключей}
             For I:=0 to SList.Count-1 do begin
                {Индекс строки}
                SRow := Trim(SList[I]);
                {Проверка на повтор}
                If SList2.IndexOf(SRow) >=0 then Continue;
                {Заголовок отчета}
                SCaption := RowsNumericToCaption(SRow, NData);
                If SCaption='' then Continue;
                {Добавляем узел}
                N := TreeForm.Items.AddChildObject(Node, SCaption, Pointer(StrToInt(SRow)));
                N.ImageIndex    := ICO_FORM_DETAIL_ROW;
                N.SelectedIndex := ICO_FORM_DETAIL_ROW+1;
                SList2.Add(SRow);
             end;
          finally
             SList2.Free;
             SList.Free;
          end;
       end;


       {***********************************************************************}
       {***********   Выделена строка: формируем список столбцов   ************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_ROW  then begin
          {Инициализация}
          SList  := TStringList.Create;
          SList2 := TStringList.Create;
          try
             {Читаем из dat-файла ключи выбранной таблицы (секции)}
             If Not GetTabKeyYMR(@SList, YMR, IntToStr(Integer(Node.Parent.Data))) then Exit;

             {Обрабатываем список}
             S:=IntToStr(NData);
             For I:=0 to SList.Count-1 do begin
                SRow := SList[I];
                SCol := Format('%3s', [TokChar(SRow, ' ')]);
                If S=SRow then If SList2.IndexOf(SCol)<0 then SList2.Add(SCol);
             end;
             SList2.Sort;

             FBASE := TBASE.Create;
             try
                {Инициализация}
                ITab := Integer(Node.Parent.Data);
                {Просматриваем список ключей}
                For I:=0 to SList2.Count-1 do begin
                   {Разделяем ключ на строку и столбец}
                   SCol := Trim(SList2[I]);
                   {Заголовок отчета}
                   SCaption := ColsNumericToCaption(SCol, ITab);
                   If SCaption='' then Continue;
                   SRow := IntToStr(NData);
                   S    := IntToStr(Integer(Node.Parent.Data));
                   S    := FloatToStrF(FBASE.GetValYMR(YMR, S, SCol, SRow), ffNumber, 18,0); //FloatToStr(FBASE.GetVal(YMR, S, SCol, SRow));
                   SCaption:=SCaption+CH_VAL+S;
                   {Добавляем узел}
                   N:=TreeForm.Items.AddChildObject(Node, SCaption, Pointer(StrToInt(SCol)));
                   N.ImageIndex    := ICO_FORM_DETAIL_COL;
                   N.SelectedIndex := ICO_FORM_DETAIL_COL+1;
                end;
             finally
                FBASE.Free;
             end;
          finally
             SList2.Free;
             SList.Free;
          end;
       end;

       {***********************************************************************}
       {*****************   Выделен столбец: читаем значение   ****************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_COL  then begin
          SCol := IntToStr(NData);
          SRow := IntToStr(Integer(Node.Parent.Data));
          ITab := Integer(Node.Parent.Parent.Data);
          StatusBar.Panels[STATUS_KOD].Text := GeneratKod(IntToStr(ITab), SRow, SCol);
       end;

       {Восстанавливаем выделение}
       LoadTreeSelect(@TreeForm, INI_SELECT, INI_SELECT_FORM);

       {Обрабатываем выделение рекурсивным вызовом}
       N:=TreeForm.Selected;
       If N<>nil then begin
          If Node=nil then TreeFormChange(nil, N)
                      else If N.AbsoluteIndex<>Node.AbsoluteIndex then TreeFormChange(nil, N);
       end;

    finally
       {Автопрокрутка окна}
       AutoScrollTree(@TreeForm);

       {Доступность Action}
       If Sender<>nil then EnablAction;

       {Возвращаем стандартный режим}
       TreeForm.Items.EndUpdate;
       TreeForm.OnChange:=TreeFormChange;
    end;
end;


{==============================================================================}
{========================    СОБЫТИЕ: OnKeyDown    ============================}
{==============================================================================}
procedure TFMAIN.TreeFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    Case Key of
    13: begin {Enter}
           If TreeForm.Selected=nil then Exit;
           If TreeForm.Selected.ImageIndex=0 then TreeForm.Selected.Expanded:=not TreeForm.Selected.Expanded
                                             else TreeFormDblClick(nil);
        end;
    end;
end;


{==============================================================================}
{=========================  СОБЫТИЕ: ON_DblClick  =============================}
{==============================================================================}
procedure TFMAIN.TreeFormDblClick(Sender: TObject);
var N : TTreeNode;
    P : TPoint;
begin
    {Инициализация}
    If TreeForm.Selected=nil then Exit;

    {Реакция только при нахождении курсора в области выбора}
    If Sender<>nil then begin
       If GetCursorPos(P) = false then Exit;
       P := TreeForm.ScreenToClient(P);
       N := TreeForm.GetNodeAt(P.X, P.Y);
    end else begin
       If TreeForm.SelectionCount=0 then Exit;
       N:=TreeForm.Selected;
    end;
    If N=nil then Exit;

    {Открытие/редактирование формы}
    AOpenExecute(Sender);
end;


{==============================================================================}
{===================    СОБЫТИЕ: OnCustomDrawItem    ==========================}
{==============================================================================}
procedure TFMAIN.TreeFormCustomDrawItem(Sender: TCustomTreeView;
                 Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
    If Node.Selected then begin
       If WIN_VER < wvVista then begin
          {Для WinXP}
          Sender.Canvas.Brush.Color := clGreen;
          Sender.Canvas.Font.Color  := clWhite;
       end else begin
          {Для WinVista и Win7}
          Sender.Canvas.Font.Color := clRed;
       end;
    end else begin
       Case Node.ImageIndex of
           // ICO_FORM_DETAIL_TABLE      : Sender.Canvas.Font.Style := Sender.Canvas.Font.Style + [fsBold];
           ICO_FORM_DETAIL_ROW,
           ICO_FORM_DETAIL_ANALITIC1,
           ICO_FORM_DETAIL_ANALITIC2  : Sender.Canvas.Font.Color := clBlue;
           ICO_FORM_DETAIL_COL        : Sender.Canvas.Font.Color := COLOR_ROW; //clNavy; //clSkyBlue;
       end;
    end;   
end;


procedure TFMAIN.TreeFormExpanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean);
begin
    {Доступность Action}
   // EnablAction;
end;

procedure TFMAIN.TreeFormCollapsing(Sender: TObject; Node: TTreeNode; var AllowCollapse: Boolean);
begin
    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{===================   ПОЛУЧЕНИЕ КОДА ВЫДЕЛЕНИЯ В TREEFORM   ==================}
{==============================================================================}
function TFMAIN.KeySelForm: String;
var NForm : TTreeNode;
begin
    {Инициализация}
    Result := '';
    NForm  := TreeForm.Selected;
    If NForm=nil then Exit;
    If NForm.ImageIndex <> ICO_FORM_DETAIL_COL then Exit;
    Result := IntToStr(Integer(NForm.Parent.Parent.Data))+CH_SPR+
              IntToStr(Integer(NForm.Data))+CH_SPR+
              IntToStr(Integer(NForm.Parent.Data));
end;


{==============================================================================}
{======================   ОЧИСТКА СТРУКТУРЫ SELFORM   =========================}
{==============================================================================}
procedure TFMAIN.ClearSelForm;
begin
    With SelForm do begin
       Node := nil;
    end;
end;


{==============================================================================}
{=====================   ЗАПОЛНЕНИЕ СТРУКТУРЫ SELFORM   =======================}
{==============================================================================}
function TFMAIN.SetSelForm: Boolean;
begin
    {Инициализация}
    Result := false;
    ClearSelForm;

    {Выделенный узел}
    SelForm.Node := TreeForm.Selected;
    If SelForm.Node=nil then Exit;

    // SKey

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{======================   ДОСТУПНОСТЬ КНОПОК ПОИСКА   =========================}
{==============================================================================}
procedure TFMAIN.CBFindChange(Sender: TObject);
begin
    EnablAction;
end;


{==============================================================================}
{======================   ПОИСК ПРЕДЫДУЩЕГО ЗНАЧЕНИЯ   ========================}
{==============================================================================}
procedure TFMAIN.AFindPrevExecute(Sender: TObject);
begin
    Finder(false);
end;

{==============================================================================}
{=======================   ПОИСК СЛЕДУЮЩЕГО ЗНАЧЕНИЯ   ========================}
{==============================================================================}
procedure TFMAIN.AFindNextExecute(Sender: TObject);
begin
    Finder(true);
end;

procedure TFMAIN.CBFindKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    If Key=13 then AFindNext.Execute;
end;


{==============================================================================}
{==============   СОБЫТИЕ: ВКЛЮЧЕНИЕ РЕЖИМА ПЕРЕТАСКИВАНИЯ   ==================}
{==============================================================================}
procedure TFMAIN.TreeFormMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
    {Инициализация}
    If Button <> mbLeft        then Exit;
    If TreeForm.Selected = nil then Exit;

    {Перетаскиваем только строки или столбцы}
    Case TreeForm.Selected.ImageIndex of
    ICO_FORM_DETAIL_ROW, ICO_FORM_DETAIL_COL: TreeForm.BeginDrag(true, 0);
    end;
end;



