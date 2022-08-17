{==============================================================================}
{=========================   ДОСТУПНОСТЬ ACTION   =============================}
{==============================================================================}
procedure TFMATRIX.EnablAction;
var BRun, BOpen, BClear : Boolean;
    STab, SCol, SRow    : String;
begin
    {Инициализация}
    BRun   := true;
    BOpen  := true;
    BClear := true;
    try
       {Таблица должна быть известна}
       If EBlockKod.Text = '' then BRun := false;

       {Должен быть блок ввода}
       If EBlockIn.Text = '' then BRun := false;
       
       {Не может быть больше одного блока ввода}
       If Pos(',', EBlockIn.Text) > 0 then BRun := false;
       If Pos(' ', EBlockIn.Text) > 0 then BRun := false;

       {Должен быть блок вывода}
       If EBlockOut.Text = '' then BRun := false;

       {Код должен быть cформирован}
       If Not SeparatKod(EBlockKod.Text, STab, SRow, SCol) then BRun:=false;
       If (STab='') or (SRow='') or (SCol='') then BRun:=false;

    finally
       ARun.Enabled   := BRun;
       AOpen.Enabled  := BOpen;
       AClear.Enabled := BClear;
    end;
end;


{==============================================================================}
{===========================   ACTION: ВЫПОЛНИТЬ   ============================}
{==============================================================================}
procedure TFMATRIX.ARunExecute(Sender: TObject);
var F                          : TIniFile;
    SList                      : TStringList;
    STab, SRow, SCol, SPage, S : String;
    FirstRow, FirstCol         : Integer;

    {==========================================================================}
    {=====   Индексация блока   ===============================================}
    {==========================================================================}
    function IndexBlock(const SBlock: String): Boolean;
    var SBlock1, SBlock2 : String;
        SKey, SVal       : String;
        Row1, Col1       : Integer;
        Row2, Col2       : Integer;
        IRow, ICol       : Integer;
    begin
        {Инициализация}
        Result:=false;

        {Начало и конец области}
        SBlock1 := AnsiUpperCase(CutSlovo(SBlock, 1, ':'));
        SBlock2 := AnsiUpperCase(CutSlovo(SBlock, 2, ':'));
        If SBlock1='' then Exit;
        If SBlock2='' then SBlock2:=SBlock1;

        {Определяем область в формате A1}
        If Not ExcelSeparatAddress(SBlock1, Col1, Row1) then Exit;
        If Not ExcelSeparatAddress(SBlock2, Col2, Row2) then Exit;
        If (Col1>Col2) or (Row1>Row2) then Exit;

        {Просматриваем все ячейки области}
        For ICol := Col1 to Col2 do begin
           For IRow := Row1 to Row2 do begin
              {Переводим адрес ячейки в форму A1}
              SKey := ExcelCombineAddress(ICol, IRow);
              If SKey='' then Exit;
              {Формируем код}
              SVal := STab+CH_SPR+IntToStr(ICol-Col1+FirstCol)+CH_SPR+IntToStr(IRow-Row1+FirstRow);
              {Записываем код}
              F.WriteString(INI_PAGE+SPage, SKey, SVal);
           end;
        end;

        {Возвращаемый результат}
        Result:=true;
    end;

begin
    {Инициализация}
    If Not SeparatKod(EBlockKod.Text, STab, SRow, SCol) then Exit;
    FirstRow := StrToInt(SRow);
    FirstCol := StrToInt(SCol);
    SPage    := EPage.Text;

    F     := TIniFile.Create(SPathIni);
    SList := TStringList.Create;
    try
       {Запоминаем индекс таблицы}
       S:=F.ReadString(INI_PARAM, INI_PARAM_TABLES, '');
       SeparatMStr(S, @SList, ', ');
       If SList.IndexOf(STab)<0 then begin
          If S<>'' then S:=S+', ';
          S:=S+STab;
          F.WriteString(INI_PARAM, INI_PARAM_TABLES, S);
       end;

       {Запоминаем индекс листа}
       S:=F.ReadString(INI_PARAM, INI_PARAM_PAGES, '');
       SeparatMStr(S, @SList, ', ');
       If SList.IndexOf(SPage)<0 then begin
          If S<>'' then S:=S+', ';
          S:=S+SPage;
          F.WriteString(INI_PARAM, INI_PARAM_PAGES, S);
       end;

        {Запоминаем область ввода в БД}
        If EBlockIn.Text<>'' then begin
           S:=F.ReadString(INI_PAGE+SPage, INI_PAGE_IN, '');
           SeparatMStr(S, @SList, ', ');
           If SList.IndexOf(EBlockIn.Text)<0 then begin
              If S<>'' then S:=S+', ';
              S:=S+EBlockIn.Text;
              F.WriteString(INI_PAGE+SPage, INI_PAGE_IN, S);
           end;
        end;

        {Запоминаем область вывода из БД}
        If EBlockOut.Text<>'' then begin
           S:=F.ReadString(INI_PAGE+SPage, INI_PAGE_OUT, '');
           SeparatMStr(S, @SList, ', ');
           If SList.IndexOf(EBlockOut.Text)<0 then begin
              If S<>'' then S:=S+', ';
              S:=S+EBlockOut.Text;
              F.WriteString(INI_PAGE+SPage, INI_PAGE_OUT, S);
           end;
       end;
       
       {Индексация области ввода}
       If Not IndexBlock(EBlockIn.Text) then Exit;

       {Информация}
       ShowMessage('Матрица скорректирована: '+CH_NEW+ExtractFileName(SPathIni));

       {Очистка полей}
       EBlockIn.Text  := '';
       EBlockOut.Text := '';
       EnablAction;

       {Установка фокуса ввода}
       If EBlockIn.Enabled then EBlockIn.SetFocus;
    finally
       If SList<>nil then SList.Free;
       F.Free;
    end;
end;


{==============================================================================}
{============================   ACTION: ОЧИСТИТЬ   ============================}
{==============================================================================}
procedure TFMATRIX.AClearExecute(Sender: TObject);
var F        : TIniFile;
    Section  : String;
begin
    {Инициализация}
    Section:=INI_PAGE+EPage.Text;
    F:=TIniFile.Create(SPathIni);
    try
       If Not F.SectionExists(Section) then begin
          MessageDlg('Индексы для страницы № '+EPage.Text+' не обнаружены!', mtError, [mbOk], 0);
          Exit;
       end;

       If MessageDlg('Подтвердите очистку индексов для листа № '+EPage.Text+'.',
       mtWarning, [mbYes, mbNo], 0)<>mrYes then Exit;

       F.EraseSection(Section);
       MessageDlg('Индексы для страницы № '+EPage.Text+' удалены!', mtInformation, [mbOk], 0);
    finally
       F.Free;
    end;
end;


{==============================================================================}
{============================   ACTION: ОТКРЫТЬ   =============================}
{==============================================================================}
procedure TFMATRIX.AOpenExecute(Sender: TObject);
begin
    StartAssociatedExe(SPathINI, SW_SHOWNORMAL);
end;


{==============================================================================}
{=====================  ОЧИСТИТЬ ПОЛЕ ИНДЕКСОВ ВВОДА    =======================}
{==============================================================================}
procedure TFMATRIX.EBlockOutButtonClick(Sender: TObject);
begin
    EBlockOut.Text := '';
    SOut           := '';
    EChange(Sender);
end;


{==============================================================================}
{===================  КОПИРОВАТЬ В ПОЛЕ ИНДЕКСОВ ВВОДА    =====================}
{==============================================================================}
procedure TFMATRIX.EBlockInAltButtonClick(Sender: TObject);
begin
    EBlockOut.Text := EBlockIn.Text;
    SOut           := '';
    EChange(Sender);
end;


{==============================================================================}
{==========  СИНХРОНИЗИРОВАТЬ ИЗМЕНЕНИЯ БЛОКА ВВОДА С БЛОКОМ ВЫВОДА  ==========}
{==============================================================================}
procedure TFMATRIX.EBlockInButtonClick(Sender: TObject);
begin
    EChange(Sender);
end;


