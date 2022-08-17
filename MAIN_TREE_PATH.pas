{******************************************************************************}
{**************************  ������ � TREE_PATH  ******************************}
{******************************************************************************}
{**********   N.Data = 0  - �������                                  **********}
{**********   N.Data = -1 - ������������� ��������                   **********}
{**********   N.Data > 0  - �������� ������                          **********}
{******************************************************************************}


{==============================================================================}
{========================   ��������� TREEPATH   ==============================}
{==============================================================================}
procedure TFMAIN.RefreshTreePath;

    {==========================================================================}
    {==================  ���������� ������� ������ ���  =======================}
    {==========================================================================}
    function SetTreeYear: Boolean;
    var N     : TTreeNode;
        LFind : array of String;
        I     : Integer;
    begin
        {�������������}
        Result:=false;
        TreePath.Items.Clear;
        try
           {����� ���}
           If FindDataYMR(@LFind, '*', false) > 0 then begin
              For I:=Low(LFind) to High(LFind) do begin
                 If IsIntegerStr(LFind[I]) then begin
                    N:=TreePath.Items.AddChildObject(nil, LFind[I], Pointer(0));
                    N.ImageIndex    := 0;
                    N.SelectedIndex := 1;
                 end;
              end;

              {��������������� ���������}
              LoadTreeSelect(@TreePath, INI_SELECT, INI_SELECT_PATH);

              {������������ ���������}
              N:=TreePath.Selected;
              TreePathChange(nil, N);

              {���������� �����}
              TreePath.AlphaSort(true);

              {������������ ���������}
              Result:=true;
           end;
        finally
           SetLength(LFind, 0);
        end;
    end;

{==============================================================================}
begin
    {�������������}
    TreePath.Items.BeginUpdate;
    TreePath.OnChange:=nil;
    try
       TreePath.Items.Clear;
       {������������� ������� ������ ���}
       If Not SetTreeYear then begin
          EnablAction;
          Exit;
       end;
       {��������� ������� �������� ������ ����}
       RefreshTreeForm;
    finally
       TreePath.OnChange := TreePathChange;
       TreePath.Items.EndUpdate;
    end;
end;


{==============================================================================}
{==========================  �������: ON_CHANGE  ==============================}
{==============================================================================}
procedure TFMAIN.TreePathChange(Sender: TObject; Node: TTreeNode);
var NBase: TTreeNode;

    {==========================================================================}
    {======================  ���������� ������ �������  =======================}
    {==========================================================================}
    function SetTreeMonth(const PNParent: PTreeNode): Boolean;
    var N     : TTreeNode;
        LFind : array of String;
        SPath : String;
        I     : Integer;
    begin
        {�������������}
        Result:=false;
        If PNParent=nil then Exit;
        If PNParent^.Count>0 then Exit;
        SPath:=GetNodePath(PNParent);
        TreePath.Items.BeginUpdate;
        try
           {����� �������}
           If FindDataYMR(@LFind, SPath+'*', false) > 0 then begin
              For I:=Low(LFind) to High(LFind) do begin
                 If MonthStrToInd(LFind[I])>0  then begin
                    N:=TreePath.Items.AddChildObject(PNParent^, LFind[I], Pointer(0));   // Pointer(IMAIN_REGION) <-- 0 1 200 - ������ �������� �������
                    N.ImageIndex    := 0;
                    N.SelectedIndex := 1;
                 end;
              end;

              {��������������� ���������}
              LoadTreeSelect(@TreePath, INI_SELECT, INI_SELECT_PATH);

              {������������ ��������� ����������� �������}
              N:=TreePath.Selected;
              If N<>nil then begin
                 If N.AbsoluteIndex<>Node.AbsoluteIndex then TreePathChange(nil, N);
              end;

              {������������ ���������}
              Result:=true;
           end;
        finally
           SetLength(LFind, 0);
           TreePath.Items.EndUpdate;
        end;
    end;


    {==========================================================================}
    {====================  ���������� ������ ��������   =======================}
    {==========================================================================}
    function SetTreeRegion(const PNParent: PTreeNode; const ILevel: Integer): Boolean;
    var N        : TTreeNode;
        LFind    : array of String;
        IFind    : Integer;
        LRegion_ : TLRegion;
        FDate    : TDate;
        S, SPath : String;
        S1, S2   : String;
        I        : Integer;
    begin
        {�������������}
        Result:=false; N:=nil;
        If PNParent=nil then Exit;
        If PNParent^.Count>0 then Exit;
        If Integer(PNParent^.Data)<0 then Exit;
        SetSelPath;          //  ����������� �.�. ���� ����� ���� ��������� �� �����������: If Not SetSelPath then Exit;
        FDate := SDatePeriod(SelPath.SYear, SelPath.SMonth);
        SPath := GetNodePath(PNParent);
        For I:=0 to ILevel do TokCharEnd(SPath, '\'); I:=0;
        If SPath<>'' then SPath:=SPath+'\';
        TreePath.Items.BeginUpdate;
        try

           {*******************************************************************}
           {*****   ������������� ������� ������ ���� �� �����  ***************}
           {*******************************************************************}
           If (ILevel=0) and (IMAIN_REGION<>0) then begin
              S := RegionsCounterToCaption(IMAIN_REGION);
              If S='' then Exit;
              If FindDataYMR(nil, SPath+S+'.dat', false) = 0 then Exit;
              N := TreePath.Items.AddChildObject(PNParent^, S, Pointer(IMAIN_REGION));
              I := ICO_PATH_MIN + (ILevel * 2);
              If I>ICO_PATH_MAX then I := ICO_PATH_MAX;

              {��������� ����}
              N.ImageIndex    := I;
              N.SelectedIndex := N.ImageIndex + 1;

              {��������������� ���������}
              LoadTreeSelect(@TreePath, INI_SELECT, INI_SELECT_PATH);

              {������������ ��������� ����������� �������}
              N:=TreePath.Selected;
              If N<>nil then begin
                 If N.AbsoluteIndex<>Node.AbsoluteIndex then TreePathChange(nil, N);
              end;

              {������������ ���������}
              Result:=true;


           {*******************************************************************}
           {*****   � ���� �������   ******************************************}
           {*******************************************************************}
           end else begin
              {����� ��������}
              try
                 If FindDataYMR(@LFind, SPath+'*', false) > 0 then begin
                    For IFind:=Low(LFind) to High(LFind) do begin
                       {�������� ��� �������}
                       S  := LFind[IFind];
                       Delete(S, Length(S)+1-4, 4);
                       S2 := CutModulStr (S,     SUB_RGN1, SUB_RGN2);
                       S1 := ReplModulStr(S, '', SUB_RGN1, SUB_RGN2);

                       {�������� ������}
                       If S2='' then begin
                          If Not FindRegions(@LRegion_, true, -1, S1, Integer(PNParent^.Data), -1, FDate) then Continue;
                          N := TreePath.Items.AddChildObject(PNParent^, S, Pointer(LRegion_[0].FCounter));
                          I := ICO_PATH_MIN + (ILevel * 2);
                          If I>ICO_PATH_MAX then I := ICO_PATH_MAX;
                       {���������}
                       end else begin
                          If PNParent^.Text<>S1 then Continue;
                          N := TreePath.Items.AddChildObject(PNParent^, S, Pointer(-1));
                          I := ICO_PATH_MAX;
                       end;

                       {��������� ����}
                       N.ImageIndex    := I;
                       N.SelectedIndex := N.ImageIndex + 1;
                    end;

                    {��������������� ���������}
                    LoadTreeSelect(@TreePath, INI_SELECT, INI_SELECT_PATH);

                    {������������ ��������� ����������� �������}
                    N:=TreePath.Selected;
                    If N<>nil then begin
                       If N.AbsoluteIndex<>Node.AbsoluteIndex then TreePathChange(nil, N);
                    end;

                    {������������ ���������}
                    Result:=true;
                 end;
              finally
                 SetLength(LFind, 0);
              end;
           end;
        finally
           SetLength(LRegion_, 0);
           TreePath.Items.EndUpdate;
        end;
    end;


{==============================================================================}
begin
    {�������������}
    ClearSelPath;

    {���������� ����}
    NBase:=TreePath.Selected;
    If NBase=nil then Exit;

    {��������� ���������}
    If Sender<>nil then SaveTreeSelect(@TreePath, INI_SELECT, INI_SELECT_PATH);

    {������� ������}
    TreePath.OnChange:=nil;
    If NBase.Level=0 then SetTreeMonth(@NBase)
                     else SetTreeRegion(@NBase, NBase.Level-1);
    TreePath.OnChange:=TreePathChange;

    If Sender<>nil then begin
       {����������}
       If NBase.Count>1 then StatusBar.Panels[STATUS_MAIN].Text := ' '+NBase.Text+' �������� '+IntToStr(NBase.Count)+' ����.'
                        else StatusBar.Panels[STATUS_MAIN].Text := '';
       {���������� �����}
       TreePath.AlphaSort(true);

       {������������� ����}
       AutoScrollTree(@TreePath);

       {��������� ������� �������� ������ ����}
       RefreshTreeForm;
    end;
end;


{==============================================================================}
{========================  �������: ON_DBL_CLICK  =============================}
{==============================================================================}
procedure TFMAIN.TreePathDblClick(Sender: TObject);
begin
    AOpenExecute(Sender);
end;


{==============================================================================}
{======================   ������� ��������� SELPATH   =========================}
{==============================================================================}
procedure TFMAIN.ClearSelPath;
begin
    With SelPath do begin
       Node      := nil;
       SPath     := '';
       IYear     := 0;
       IMonth    := 0;
       SYear     := '';
       SMonth    := '';
       SRegion   := '';
       Date      := 0;
       DateBegin := 0;
       DateEnd   := 0;
    end;
end;


{==============================================================================}
{=====================   ���������� ��������� SELPATH   =======================}
{==============================================================================}
function TFMAIN.SetSelPath: Boolean;
label Er;
begin
    {�������������}
    Result := false;
    ClearSelPath;

    {���������� ����}
    SelPath.Node := TreePath.Selected;
    If SelPath.Node=nil then Exit;

    {���������� ������ � ������}
    SelPath.SPath   := GetNodePath(@SelPath.Node);   // 2009\����\���������� ��������\�����������\�������������\
    If GetColSlov(SelPath.SPath, '\')<3 then Goto Er;
    SelPath.SYear   := CutSlovo(SelPath.SPath, 1, '\');
    SelPath.SMonth  := CutSlovo(SelPath.SPath, 2, '\');
    SelPath.SRegion := CutSlovoEndChar(SelPath.SPath, 1, '\');
    If SelPath.SRegion='' then Goto Er;
    If (Not IsIntegerStr(SelPath.SYear)) then Goto Er;
    SelPath.IYear   := StrToInt(SelPath.SYear);
    SelPath.IMonth  := MonthStrToInd(SelPath.SMonth);
    If SelPath.IMonth=0 then Goto Er;

    {���������� ����}
    SelPath.Date      := IDatePeriod(SelPath.IYear, SelPath.IMonth);
    SelPath.DateBegin := EncodeDate(SelPath.IYear, SelPath.IMonth, 1);
    SelPath.DateEnd   := DateTimeCorrect(SelPath.DateBegin, 0, 1, 0, 0, 0);
    If (SelPath.Date=0) or (SelPath.DateBegin=0) or (SelPath.DateEnd=0) then Goto Er;

    {������������ ���������}
    Result:=true;
    Exit;

Er: ClearSelPath;
end;


{==============================================================================}
{=======================    �������� ����������    ============================}
{==============================================================================}
procedure TFMAIN.TreePathCompare(Sender: TObject; Node1, Node2: TTreeNode;
                                 Data: Integer; var Compare: Integer);
var I1, I2: Word;
begin
    If (Node1=nil) or (Node2=nil) then Exit;
    If (Node1.Level=1) and (Node2.Level=1) then begin
       I1:=MonthStrToInd(Node1.Text);
       I2:=MonthStrToInd(Node2.Text);
       If I1<I2 then Compare:=-1 else Compare:=1;
       Exit;
    end;
    If Node1.Text < Node2.Text then Compare:=-1 else Compare:=1;
end;


{==============================================================================}
{===================    �������: OnCustomDrawItem    ==========================}
{==============================================================================}
procedure TFMAIN.TreePathCustomDrawItem(Sender: TCustomTreeView;
          Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
    If Not Node.Selected then Exit;
    If WIN_VER < wvVista then begin
       {��� WinXP}
       Sender.Canvas.Brush.Color := clGreen;
       Sender.Canvas.Font.Color  := clWhite;
    end else begin
       {��� WinVista � Win7}
       Sender.Canvas.Font.Color := clRed;
    end;
end;


{==============================================================================}
{======================    �������: OnResizeNav    ============================}
{==============================================================================}
procedure TFMAIN.PPathNavResize(Sender: TObject);
var I: Integer;
begin
    I := (PPathNav.ClientWidth Div 6) - 1;
    BtnPrev12.Width := I;
    BtnPrev3.Width  := I;
    BtnPrev1.Width  := I;
    BtnNext1.Width  := I;
    BtnNext3.Width  := I;
    BtnNext12.Width := I;

    BtnPrev3.Left  := I + 1;
    BtnPrev1.Left  := (I * 2) + 1;
    BtnNext3.Left  := (I * 2) + 1;
    BtnNext1.Left  := (I * 2) + 1;
end;


{==============================================================================}
{====================    ACTION: �������� ������    ===========================}
{==============================================================================}
procedure TFMAIN.ANavPathExecute(Sender: TObject);
var N             : TTreeNode;
    IYear, IMonth : Integer;
    S, SPath      : String;
begin
    {�������������}
    N := TreePath.Selected;

    {���������� ���������� ������}
    SPath  := GetNodePath(@N);    // 2009\����\���������� ��������\�����������\�������������\
    S      := TokChar(SPath, '\');
    If Not IsIntegerStr(S) then Exit;
    IYear  := StrToInt(S);
    S      := TokChar(SPath, '\');
    IMonth := MonthStrToInd(S);
    If IMonth=0 then Exit;

    {������������ ����}
    IMonth := IMonth + TAction(Sender).Tag;
    If IMonth < 1 then begin IYear:=IYear-1; IMonth:=IMonth+12; end;
    If IMonth >12 then begin IYear:=IYear+1; IMonth:=IMonth-12; end;
    SPath:=IntToStr(IYear)+'\'+MonthIndToStr(IMonth)+'\'+SPath;

    {��������� ����� ������}
    WriteLocalString(INI_SELECT, INI_SELECT_PATH, SPath);

    {������������ TreePath}
    RefreshTreePath;
    //TreePathChange(nil, Node);
end;


