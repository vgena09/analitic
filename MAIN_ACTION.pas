{******************************************************************************}
{**********************  К О М А Н Д Ы    A C T I O N  ************************}
{******************************************************************************}

            
{==============================================================================}
{===========================   ДОСТУПНОСТЬ ACTION    ==========================}
{==============================================================================}
procedure TFMAIN.EnablAction;
var IsOk, IsReport, IsAnal, IsTab, IsRow, IsCol : Boolean;
    BParam, BSubRgn, BMainRgn : Boolean;
    ILevel : Integer;
    N      : TTreeNode;
    SPath  : String;
    IndImg : Integer;
begin
    {*** Инициализация TreePath ***********************************************}
    N := TreePath.Selected;
    If N <> nil then begin
       SPath    := GetNodePath(@N);
       ILevel   := N.Level;
       BMainRgn := CmpStr(N.Text, SMAIN_REGION);
       BSubRgn  := (FindStr('[', N.Text) > 0);
    end else begin
       SPath    := '';
       ILevel   := -1;
       BMainRgn := false;
       BSubRgn  := false;
    end;

    {Признак наличия параметров аналитики}
    If FFPARAM <> nil then BParam := Length(TFPARAM(FFPARAM).LParam) > 0
                      else BParam := false;

    IsOk                   := (Not IsBlock) and (ILevel > 1);
    ANavPrev12.Enabled     := IsOk;
    ANavPrev3.Enabled      := IsOk;
    ANavPrev1.Enabled      := IsOk;
    ANavNext1.Enabled      := IsOk;
    ANavNext3.Enabled      := IsOk;
    ANavNext12.Enabled     := IsOk;

    AAnalizRegion.Enabled  := (Not IsBlock) and (ILevel > 1) and BParam and (Not BMainRgn);
    AAnalizCompare.Enabled := (Not IsBlock) and (ILevel > 1) and BParam and (Not BSubRgn);
    AAnalizRating.Enabled  := (Not IsBlock) and (ILevel > 1) and BParam and (Not BSubRgn);
    AAnalizDynamic.Enabled := (Not IsBlock) and (ILevel > 1) and BParam;

    ADelDir.Enabled        := (Not IsBlock) and (ILevel >=0 ) and (ILevel < 2);

    {*** Инициализация TreeForm ***********************************************}
    If TreeForm.Selected <> nil then IndImg := TreeForm.Selected.ImageIndex
                                else IndImg := -1;
    //IsOk     := (Not IsBlock) and (IndImg >= 0);
    IsTab    := (Not IsBlock) and (IndImg = ICO_FORM_DETAIL_TABLE);
    IsRow    := (Not IsBlock) and (IndImg = ICO_FORM_DETAIL_ROW);
    IsCol    := (Not IsBlock) and (IndImg = ICO_FORM_DETAIL_COL);
    IsReport := (Not IsBlock) and (IndImg = ICO_FORM_DETAIL_REPORT);
    IsAnal   := (Not IsBlock) and (IndImg > 0) and (Not IsTab) and (Not IsCol) and (Not IsRow);

    ABlank.Enabled         := Not IsBlock;
    ANew.Enabled           := (Not IsBlock) and IS_ADMIN;
    AImport.Enabled        := (Not IsBlock) and IS_ADMIN;
    AExport.Enabled        := (Not IsBlock) and IS_ADMIN;

    AOpen.Enabled          := (Not IsBlock) and (IsReport or IsCol or IsAnal);

    BtnServ.Enabled        := Not IsBlock;
    AHelp.Enabled          := Not IsBlock;
    AClose.Enabled         := Not IsBlock;

    CBFind.Enabled         := Not IsBlock;
    AFindPrev.Enabled      := Not IsBlock;
    AFindNext.Enabled      := Not IsBlock;

    AVerify.Enabled        := IsReport and IS_ADMIN;;
    AMatrixIn.Enabled      := IsReport and IS_ADMIN;;
    AMatrixSet.Enabled     := IsReport and IS_ADMIN;;
    ADel.Enabled           := IsReport and IS_ADMIN;;

    AMatrixOut.Enabled     := (IsReport or IsAnal) and IS_ADMIN;;
    AFormXLS.Enabled       := (IsReport or IsAnal) and IS_ADMIN;;
    AFormINI.Enabled       := (IsReport or IsAnal) and IS_ADMIN;;

    AKodCorrect.Enabled    := IS_ADMIN; //(IsTab or IsCol or IsRow) and IS_ADMIN;
end;


{==============================================================================}
{===========================   ACTION: АНАЛИЗ   ===============================}
{==============================================================================}
procedure TFMAIN.AAnalizExecute(Sender: TObject);
var F_REPORT : TREPORT;
    ID       : Integer;
    B        : Boolean;
begin
    {Инициализация}
    ID := TAction(Sender).Tag;
    If ID = 0 then Exit;
    If Not SetSelPath then Exit;

    {Контроль не нужен}
    IsControl:=false;

    {Заполнение отчета}
    B:=false;
    SetBlock;
    try
       F_REPORT := TREPORT.Create;
       try
          {Создаем отчет}
          Case ID of
             TAG_ANALIZ_COMPARE: B := F_REPORT.CreateReport(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '', IDFORM_ACOMPARE, 0, 0, false);
             TAG_ANALIZ_RATING:  B := F_REPORT.CreateReport(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '', IDFORM_ARATING,  0, 0, false);
             TAG_ANALIZ_REGION:  B := F_REPORT.CreateReport(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '', IDFORM_AREGION,  0, 0, false);
             TAG_ANALIZ_DYNAMIC: B := F_REPORT.CreateReport(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '', IDFORM_ADYNAMIC, 0, 0, false);
          end;
       finally
           F_REPORT.Free;
       end;
    finally
       {Режим ожидания если отчет создан}
       If B and (XLS.Workbooks.Count > 0) then begin
          IsModify := false;
          Enabled  := false;
       end else begin
          ClearBlock;
       end;
    end;
end;


{==============================================================================}
{===========================   ACTION: ЭКСПОРТ   ==============================}
{==============================================================================}
procedure TFMAIN.AExportExecute(Sender: TObject);
var F: TFEXPORT;
begin
    F:=TFEXPORT.Create(Self);
    try     F.Execute;
    finally F.Free; end;
end;


{==============================================================================}
{===========================   ACTION: ПРОВЕРКА   =============================}
{==============================================================================}
procedure TFMAIN.AVerifyExecute(Sender: TObject);
var NForm : TTreeNode;
    FIni  : TIniFile;
    LTab  : TStringList;
    SPath : String;
    S     : String;
begin
    {Определяем выделение}
    If Not SetSelPath then Exit;
    NForm:=TreeForm.Selected;
    If NForm=nil then Exit;

    {Путь к ini-файлу}
    SPath := FormsCounterToFile(Integer(NForm.Data));
    If SPath='' then begin ErrMsg('Не определен тип проверяемого файла!'); Exit; end;
    SPath := ChangeFileExt(SPath, '.ini');
    If Not FileExists(SPath) then begin ErrMsg('Не определен ini-файл проверяемого файла!'); Exit; end;

    LTab:=TStringList.Create;
    try
       {Получаем список проверяемых таблиц из ini-файла шаблона}
       FIni:=TIniFile.Create(SPath);
       try     S:=FIni.ReadString(INI_PARAM, INI_PARAM_TABLES, '');
       finally FIni.Free;
       end;
       SeparatMStr(S, @LTab, ', ');
       If LTab.Count=0 then begin ErrMsg('Не обнаружен список проверяемых таблиц!'); Exit; end;

       {Производим проверку}
       VerifyTable(@LTab, SelPath.SYear, SelPath.SMonth, SelPath.SRegion, true);
    finally
       LTab.Free;
    end;
end;


{==============================================================================}
{========================   ACTION: УДАЛИТЬ ОТЧЕТ   ===========================}
{==============================================================================}
procedure TFMAIN.ADelExecute(Sender: TObject);
var LRegion_          : TLRegion;
    NForm             : TTreeNode;
    FDat              : TMemIniFile;
    FDate             : TDate;
    S, SPath          : String;
    LTab, LReg, SList : TStringList;
    ITab, IReg, I     : Integer;
    B, IsControl      : Boolean;
begin
    {Инициализация}
    If Not IS_ADMIN   then Exit;
    If Not SetSelPath then Exit;
    FDate := SDatePeriod(SelPath.SYear, SelPath.SMonth);
    NForm := TreeForm.Selected;
    If NForm=nil then Exit;

    {Путь к ini-файлу отчета}
    SPath := FormsCounterToFile(Integer(NForm.Data));
    If SPath='' then begin ErrMsg('Не определен тип удаляемого отчета!'); Exit; end;
    SPath := ChangeFileExt(SPath, '.ini');
    If Not FileExists(SPath) then begin ErrMsg('Не определен ini-файл удаляемого отчета!'); Exit; end;

    LTab  := TStringList.Create;
    LReg  := TStringList.Create;
    SList := TStringList.Create;
    try
       {Список удаляемых таблиц}
       If Not GetListTabIni(@LTab, SPath) then begin ErrMsg('Не определен список удаляемых таблиц!'); Exit; end;
       If LTab.Count=0 then Exit;

       {Удаляемый отчет должен допускать модернизацию}
       If Not IsModifyRegion(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, @LTab) then begin
          MessageDlg('В целях обеспечения целостности данных удаление отчета заблокировано.'+CH_NEW+
                     'Удалять необходимо отчеты структурных подразделений!', mtInformation, [mbOk], 0); Exit;
       end;

       {Если отчет стандартный, то обрабатывается список регионов по восходящей}
       IsControl := Not IsTableNoStandart(LTab[0]);
       If IsControl then begin
          If Not ListRegionsRangeUp(@LReg, SelPath.SRegion, '', '') then begin ErrMsg('Не определен список регионов!'); Exit; end;
       {Если отчет нестандартный, то обрабатывается только он один}
       end else begin
          LReg.Add(SelPath.SRegion);
       end;

       {Подтверждение операции}
       StatusBar.Panels[STATUS_MAIN].Text:=' Удаление отчета: подтверждение';
       If MessageDlg('Подтвердите удаление отчета:'+CH_NEW+NForm.Text+'!', mtWarning, [mbYes, mbNo], 0)<>mrYes then Exit;
       Repaint;
       StatusBar.Panels[STATUS_MAIN].Text:=' Удаление отчета';

       {Поочередно обрабатываем регионы}
       For IReg:=0 to LReg.Count-1 do begin
          {Файл очередного региона}
          SPath := PathRegion(SelPath.SYear, SelPath.SMonth, LReg[IReg]);
          If Not FileExists(SPath) then begin ErrMsg('Не найден файл: '+SPath); Exit; end;

          {Поочередно обрабатываем таблицы}
          For ITab:=0 to LTab.Count-1 do begin

             {Доступ к файлу}
             FDat:=TMemIniFile.Create(SPath);
             try     {Удаляем секцию}
                     FDat.EraseSection(LTab[ITab]);
                     {Список оставшихся секций}
                     SList.Clear;
                     FDat.ReadSections(SList);
                     FDat.UpdateFile;
             finally FDat.Free;
             end;

             {Если в базовом отчете нет секций - удаляем его}
             If (IReg=0) and (SList.Count=0) then DeleteFile(SPath);

             {Если в родительском отчете нет секций и зависимых регионов - удаляем его}
             If IsControl then begin
                If (IReg>0) and (SList.Count=0) then begin
                   B := true;
                   try If Not FindRegions(@LRegion_, true, -1, LReg[IReg], -1, -1, FDate) then Continue;
                       FindRegions(@LRegion_, false, -1, '', LRegion_[0].FCounter, -1, FDate);
                       For I:=0 to Length(LRegion_)-1 do begin
                          S := PathRegion(SelPath.SYear, SelPath.SMonth, LRegion_[I].FCaption);
                          If FileExists(S) then begin B:=false; Break; end;
                       end;
                   finally SetLength(LRegion_, 0);
                   end;
                   If B then begin
                      DeleteFile(SPath);
                      Break;
                   end;
                end;
             end;
          end;
       end;
    finally
       SList.Free;
       LReg.Free;
       LTab.Free;
       StatusBar.Panels[STATUS_MAIN].Text:='';
       ARefresh.Execute;
    end;
end;


{==============================================================================}
{=======================   ACTION: УДАЛИТЬ КАТАЛОГ   ==========================}
{==============================================================================}
procedure TFMAIN.ADelDirExecute(Sender: TObject);
var NPath : TTreeNode;
    SPath : String;
begin
    {Инициализация}
    If Not IS_ADMIN   then Exit;
    NPath := TreePath.Selected;
    If NPath=nil then Exit;
    SPath := GetNodePath(@NPath);
    If SPath='' then Exit;

    {Формируем путь}
    ClearSelPath;
    SelPath.SYear  := CutSlovo(SPath, 1, '\');
    SelPath.SMonth := CutSlovo(SPath, 2, '\');
    If GetColSlov(SPath, '\')>2 then Exit;
    SPath:=SelPath.SYear+'\';
    If SelPath.SMonth  <> '' then SPath := SPath + SelPath.SMonth  + '\';

    {Подтверждение операции}
    If MessageDlg('Подтвердите удаление:'+CH_NEW+SPath, mtWarning, [mbYes, mbNo], 0)<>mrYes then Exit;
    Repaint;

    {Удаляем каталог}
    SPath:=PATH_BD_DATA+SPath;
    try
       StatusBar.Panels[STATUS_MAIN].Text:=' Удаление каталога';
       If Not DelDir(SPath, true) then begin ErrMsg('Ошибка удаления каталога: '+SPath); Exit; end;
       TokCharEnd(SPath, '\');
       TokCharEnd(SPath, '\');
       RemoveDirectory(PChar(SPath));
    finally
       StatusBar.Panels[STATUS_MAIN].Text:='';
       ARefreshExecute(Sender);
    end;
end;


{==============================================================================}
{=============================   ACTION: КОД   ================================}
{==============================================================================}
procedure TFMAIN.AKodExecute(Sender: TObject);
var F                : TFKOD;
    NForm            : TTreeNode;
    STab, SCol, SRow : String;
begin
    {Инициализация}
    STab:=''; SCol:=''; SRow:='';
    NForm  := TreeForm.Selected;

    {Определяем код}
    If NForm<>nil then begin
       Case NForm.ImageIndex of
       ICO_FORM_DETAIL_TABLE: begin
          STab := IntToStr(Integer(NForm.Data));
       end;
       ICO_FORM_DETAIL_ROW: begin
          STab := IntToStr(Integer(NForm.Parent.Data));
          SRow := IntToStr(Integer(NForm.Data));
       end;
       ICO_FORM_DETAIL_COL: begin
          STab := IntToStr(Integer(NForm.Parent.Parent.Data));
          SRow := IntToStr(Integer(NForm.Parent.Data));
          SCol := IntToStr(Integer(NForm.Data));
       end;
       end;
    end;

    {Диалог}
    F:=TFKOD.Create(Self);
    try     F.Execute(GeneratKod(STab, SRow, SCol), true);
    finally F.Free;
    end;

    {Обновляем дерево форм}
    RefreshTreeForm;
end;

{==============================================================================}
{=======================   ACTION: КОРРЕКЦИЯ КОДА   ===========================}
{==============================================================================}
procedure TFMAIN.AKodCorrectExecute(Sender: TObject);
var F                : TFKOD_CORRECT;
    NForm            : TTreeNode;
    STab, SCol, SRow : String;
begin
    {Инициализация}
    NForm  := TreeForm.Selected;
    If NForm=nil then Exit;

    {Определяем код}
    STab:=''; SCol:=''; SRow:='';
    Case NForm.ImageIndex of
        ICO_FORM_DETAIL_TABLE: begin
           STab := IntToStr(Integer(NForm.Data));
        end;
        ICO_FORM_DETAIL_ROW: begin
           STab := IntToStr(Integer(NForm.Parent.Data));
           SRow := IntToStr(Integer(NForm.Data));
        end;
        ICO_FORM_DETAIL_COL: begin
           STab := IntToStr(Integer(NForm.Parent.Parent.Data));
           SRow := IntToStr(Integer(NForm.Parent.Data));
           SCol := IntToStr(Integer(NForm.Data));
        end;
    end;

    {Форма коррекции кода}
    F:=TFKOD_CORRECT.Create(Self);
    try     F.Execute(STab, SCol, SRow);
    finally F.Free;
    end;

    {Обновляем дерево форм}
    RefreshTreeForm;
end;


{==============================================================================}
{==========================   ACTION: ОБНОВИТЬ   ==============================}
{==============================================================================}
procedure TFMAIN.ARefreshExecute(Sender: TObject);
begin
    RefreshTreePath;
end;


{==============================================================================}
{======================   ACTION: БЛАНК ДЛЯ РЕГИОНА   =========================}
{==============================================================================}
procedure TFMAIN.ABlankExecute(Sender: TObject);
var F: TFBLANK;
begin
    F:=TFBLANK.Create(Self);
    try     F.ShowModal;
    finally F.Free; end;
end;


{==============================================================================}
{=========================   ACTION: ИНФОРМАТОР   =============================}
{==============================================================================}
procedure TFMAIN.AInfoExecute(Sender: TObject);
begin
    StartAssociatedExe(PATH_BD+FILE_LOCAL_INFO, SW_SHOWNORMAL);
end;


{==============================================================================}
{==========================   ACTION: НАСТРОЙКА   =============================}
{==============================================================================}
procedure TFMAIN.ASetExecute(Sender: TObject);
var F: TFSET;
begin
    F:=TFSET.Create(Self);
    try     F.ShowModal;
    finally F.Free; end;
end;


{==============================================================================}
{=======================   ACTION: ОТКРЫТЬ ШАБЛОН   ===========================}
{==============================================================================}
procedure TFMAIN.AFormXLSExecute(Sender: TObject);
var NForm : TTreeNode;
    SPath : String;
begin
    {Инициализация}
    If Not IS_ADMIN   then Exit;
    If Not SetSelPath then Exit;
    NForm:=TreeForm.Selected;
    If NForm=nil then Exit;

    {Путь к файлу}
    SPath := FormsCounterToFile(Integer(NForm.Data));
    If Not FileExists(SPath) then begin ErrMsg('Не определен файл открываемого шаблона!'+CH_NEW+SPath); Exit; end;

    {Открываем файл}
    StartAssociatedExe(SPath, SW_MAXIMIZE); //SW_SHOWNORMAL);
end;


{==============================================================================}
{=======================   ACTION: ОТКРЫТЬ МАТРИЦУ   ==========================}
{==============================================================================}
procedure TFMAIN.AFormINIExecute(Sender: TObject);
var NForm : TTreeNode;
    SPath : String;
begin
    {Инициализация}
    If Not IS_ADMIN   then Exit;
    If Not SetSelPath then Exit;
    NForm:=TreeForm.Selected;
    If NForm=nil then Exit;

    {Путь к файлу}
    SPath := FormsCounterToFile(Integer(NForm.Data));
    If SPath='' then begin ErrMsg('Не определен тип открываемого файла!'); Exit; end;
    SPath := ChangeFileExt(SPath, '.ini');
    If Not FileExists(SPath) then begin ErrMsg('Не определен ini-файл открываемого шаблона!'); Exit; end;

    {Открываем файл}
    StartAssociatedExe(SPath, SW_SHOWNORMAL); 
end;


{==============================================================================}
{====================   ACTION: СОЗДАТЬ БАЗУ ДАННЫХ РЕГИОНА  ==================}
{==============================================================================}
procedure TFMAIN.ABDRegionExecute(Sender: TObject);
var F: TFBDREGION;
begin
    F:=TFBDREGION.Create(Self);
    try     F.ShowModal;
    finally F.Free; end;
end;


{==============================================================================}
{============================   ACTION: ПОМОЩЬ    =============================}
{==============================================================================}
procedure TFMAIN.AHelpExecute(Sender: TObject);
begin
    StartAssociatedExe('"'+HelpFile+'"', SW_SHOWNORMAL);
end;


{==============================================================================}
{=============================   ACTION: ВЫХОД    =============================}
{==============================================================================}
procedure TFMAIN.ACloseExecute(Sender: TObject);
begin
    Close;
end;


{==============================================================================}
{=========================   ACTION: О ПРОГРАММЕ   ============================}
{==============================================================================}
procedure TFMAIN.AAboutExecute(Sender: TObject);
var FFABOUT: TFABOUT;
begin
    FFABOUT:=TFABOUT.Create(Self);
    try     FFABOUT.ShowModal;
    finally FFABOUT.Free;
    end;
end;


{==============================================================================}
{======================   ACTION: КОНСТРУКТОР МАТРИЦЫ  ========================}
{==============================================================================}
procedure TFMAIN.AMatrixSetExecute(Sender: TObject);
var F     : TFMATRIX;
    NForm : TTreeNode;
    SPath : String;
begin
    {Инициализация}
    NForm:=TreeForm.Selected;
    If NForm=nil then Exit;

    {Файл-источник}
    SPath := FormsCounterToFile(Integer(NForm.Data));
    If SPath='' then begin ErrMsg('Не найден шаблон отчета!'); Exit; end;

    {Запускаем конструктор}
    F:=TFMATRIX.Create(Self);
    try     F.Execute(SPath);
    finally F.Free;
    end;
end;


{==============================================================================}
{=====================   ACTION: МАТРИЦА ВВОДА/ВЫВОДА    ======================}
{==============================================================================}
procedure TFMAIN.AMatrixExecute(Sender: TObject);
var F        : TREPORT;
    NForm    : TTreeNode;
    IDMatrix : Integer;
begin
    {Инициализация}
    If (Not AMatrixIn.Enabled) and (Not AMatrixOut.Enabled) then Exit;
    If Not SetSelPath then Exit;
    NForm:=TreeForm.Selected;
    If NForm=nil then Exit;
    IDMatrix:=Integer(TAction(Sender).Tag);
    If (IDMatrix<>1) and (IDMatrix<>2) then begin ErrMsg('Недопустимая команда!'); Exit; end;

    {Заполнение отчета}
    SetBlock;
    try
       F := TREPORT.Create;
       try     F.CreateReport(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '', Integer(NForm.Data), IDMatrix, 0, false);
       finally F.Free;
       end;
       WB.Disconnect;
       XLS.Disconnect;
    finally
       ClearBlock;
    end;
end;


{==============================================================================}
{=========================   ACTION: БАЗА ДАННЫХ   ============================}
{==============================================================================}
procedure TFMAIN.AOpenBaseExecute(Sender: TObject);
begin
    StartAssociatedExe('"'+PATH_BD_DATA+'"', SW_SHOWNORMAL);
end;


{==============================================================================}
{=======================   ACTION: ВРЕМЕННЫЕ ФАЙЛЫ   ==========================}
{==============================================================================}
procedure TFMAIN.AOpenTempExecute(Sender: TObject);
begin
    StartAssociatedExe('"'+PATH_WORK+'"', SW_SHOWNORMAL);
end;




