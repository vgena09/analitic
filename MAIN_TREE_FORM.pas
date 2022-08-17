{******************************************************************************}
{**************************  ������ � TREE_FORM  ******************************}
{******************************************************************************}


{==============================================================================}
{========================   ���������  TREEFORM    ============================}
{==============================================================================}
procedure TFMAIN.RefreshTreeForm;
begin
    Repaint;
    TreeForm.Items.BeginUpdate;
    try
       {������� �����}
       TreeForm.OnChange:=nil;
       TreeForm.Items.Clear;
       TreeForm.OnChange:=TreeFormChange;

       {���������� �����}
       TreeFormChange(nil, nil);
       
    finally
       TreeForm.Items.EndUpdate;
    end;

    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{==========================  �������: ON_CHANGE  ==============================}
{==============================================================================}
{========================== Sender=nil - ������� ==============================}
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
        {�������������}
        Result:=false;

        {��������� ������� ����� ����� � ini-�����}
        If Not FileExists(SForm) then Exit;
        If Not FileExists(SIni)  then Exit;

        {���������� � ��������� � TreePath}
        If Not SetSelPath then Exit;

        {�������� ������� ���, ����� � ������ �����}
        NPath:=TreePath.Selected;
        If NPath=nil then Exit;
        If NPath.Level<2 then Exit;

        {��������� ������� ���� �� ����� ������-�������}
        Result := IsTableDatIncludeIniYMR(SIni, SelPath.SYear+'\'+SelPath.SMonth+'\'+SelPath.SRegion+'.dat');
    end;

    function GetCaption(const PQ: PADOTable; const SFormula: String): String;
    var S: String;
    begin
        {�������������}
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
    {�������������}
    StatusBar.Panels[STATUS_KOD].Text := '';

    {��������� ���������}
    If Sender<>nil then SaveTreeSelect(@TreeForm, INI_SELECT, INI_SELECT_FORM);

    {�������� TreePath: ���� ��� dat-�����, �� ������� ������� ���������}
    If Not SetSelPath then Exit;
    YMR := SetYMR(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '');
    If FindDataYMR(nil, YMR, false) = 0 then Exit;

    {������������� ����������}
    TreeForm.OnChange:=nil;
    TreeForm.Items.BeginUpdate;
    try
       {���������� � ��������� � TreeForm}
       If Node <> nil then begin
          StatusBar.Panels[STATUS_MAIN].Text := ' ' + Node.Text;
          If Node.Count > 0 then Exit;    // ����, ������� ��������� ����, �� ��������������
          NData := Integer(Node.Data);
          NIco  := Node.ImageIndex;
          If NData < 0 then Exit;         // ���� � Data < 0 �� ��������������
       end else begin
          NData := 0;
          NIco  := 0;
       end;

       {***********************************************************************}
       {****  �������� ����� ��� ������ �� ��������: ��������� ������ ���� ****}
       {***********************************************************************}
       If NIco = 0 then begin
          try
             {���� �����, ��������� �� NForm}
             If Not FindForms(@LForma_, false, false, -1, NData, SelPath.SRegion, '', '', SelPath.SMonth, false, SelPath.Date) then begin
                {���� ��� ����: ������� ������ � �����}
                If Node<>nil then If Node.ImageIndex=0 then SetNodeState(@Node, TVIS_CUT);
                Exit;
             end;

             {������������� ���������� �����}
             For IForm:=Low(LForma_) to High(LForma_) do begin
                {���� � ������ �����}
                SForm := PathFileForm(LForma_[IForm].SBankPref, LForma_[IForm].FFile);
                SIni  := ChangeFileExt(SForm, '.ini');

                {���������� ��������� ����}
                SCaption := CutLongStr(LForma_[IForm].FCaption, 140);
                If SCaption='' then Continue;

                {��������� ������������ ������� ����� � ������}
                I := LForma_[IForm].FIcon;
                If (Not VerifyForm) and (I<>0) then Continue;

                {��������� ����}
                N:=TreeForm.Items.AddChildObject(Node, SCaption, Pointer(LForma_[IForm].FCounter));

                {�������� ������}
                If I>0 then begin
                   I := ICO_FORM_MIN + ((I - 1) * 2);
                   If I > ICO_FORM_MAX then I := ICO_FORM_MAX;
                end;
                N.ImageIndex    := I;
                N.SelectedIndex := N.ImageIndex+1;
             end;
          finally
             {����������� ������}
             SetLength(LForma_, 0);
          end;


          {********************************************************************}
          {***********************  ������������ �����  ***********************}
          {********************************************************************}
          If (Sender=nil) and (Node=nil) and (SelPath.SRegion<>'') then begin
             try
                If FindDataYMR(@LFind, SelPath.SPath+'*', true) > 0 then begin
                   For IFind:=Low(LFind) to High(LFind) do begin
                      S    := ExtractFileNameWithoutExt(LFind[IFind]);
                      SExt := AnsiUpperCase(ExtractFileExt(LFind[IFind]));
                      {�������� ������}
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
       {*************   ������� �����: ��������� ������ ������   **************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_REPORT  then begin
          {�������������}
          SForm := FormsCounterToFile(NData);
          SIni  := ChangeFileExt(SForm, '.ini');
          SList := TStringList.Create;
          try
             {������ �� ini-����� ������ ���������� ������}
             If Not GetListTabIni(@SList , SIni) then Exit;
             {������������� ������ ���������� ������}
             For I:=0 to SList.Count-1 do begin
                {���� ���� ���� ������� � dat-�����}
                If IsExistsRowYMR(YMR, SList[I], '') then begin
                   {��������� ������}
                   SCaption := TablesCounterToCaption(SList[I]);
                   If SCaption='' then Continue;
                   {��������� ����}
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
       {************   �������� �������: ��������� ������ �����   *************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_TABLE  then begin
          {�������������}
          SList  := TStringList.Create;
          SList2 := TStringList.Create;
          try
             {������ �� dat-����� ����� ��������� ������� (������)}
             If Not GetTabKeyYMR(@SList, YMR, IntToStr(NData)) then Exit;
             For I:=0 to SList.Count-1 do begin
                SList[I]:=Format('%3s', [CutSlovo(SList[I], 2, ' ')]);
             end;
             SList.Sort;
             {������������� ������ ������}
             For I:=0 to SList.Count-1 do begin
                {������ ������}
                SRow := Trim(SList[I]);
                {�������� �� ������}
                If SList2.IndexOf(SRow) >=0 then Continue;
                {��������� ������}
                SCaption := RowsNumericToCaption(SRow, NData);
                If SCaption='' then Continue;
                {��������� ����}
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
       {***********   �������� ������: ��������� ������ ��������   ************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_ROW  then begin
          {�������������}
          SList  := TStringList.Create;
          SList2 := TStringList.Create;
          try
             {������ �� dat-����� ����� ��������� ������� (������)}
             If Not GetTabKeyYMR(@SList, YMR, IntToStr(Integer(Node.Parent.Data))) then Exit;

             {������������ ������}
             S:=IntToStr(NData);
             For I:=0 to SList.Count-1 do begin
                SRow := SList[I];
                SCol := Format('%3s', [TokChar(SRow, ' ')]);
                If S=SRow then If SList2.IndexOf(SCol)<0 then SList2.Add(SCol);
             end;
             SList2.Sort;

             FBASE := TBASE.Create;
             try
                {�������������}
                ITab := Integer(Node.Parent.Data);
                {������������� ������ ������}
                For I:=0 to SList2.Count-1 do begin
                   {��������� ���� �� ������ � �������}
                   SCol := Trim(SList2[I]);
                   {��������� ������}
                   SCaption := ColsNumericToCaption(SCol, ITab);
                   If SCaption='' then Continue;
                   SRow := IntToStr(NData);
                   S    := IntToStr(Integer(Node.Parent.Data));
                   S    := FloatToStrF(FBASE.GetValYMR(YMR, S, SCol, SRow), ffNumber, 18,0); //FloatToStr(FBASE.GetVal(YMR, S, SCol, SRow));
                   SCaption:=SCaption+CH_VAL+S;
                   {��������� ����}
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
       {*****************   ������� �������: ������ ��������   ****************}
       {***********************************************************************}
       If NIco = ICO_FORM_DETAIL_COL  then begin
          SCol := IntToStr(NData);
          SRow := IntToStr(Integer(Node.Parent.Data));
          ITab := Integer(Node.Parent.Parent.Data);
          StatusBar.Panels[STATUS_KOD].Text := GeneratKod(IntToStr(ITab), SRow, SCol);
       end;

       {��������������� ���������}
       LoadTreeSelect(@TreeForm, INI_SELECT, INI_SELECT_FORM);

       {������������ ��������� ����������� �������}
       N:=TreeForm.Selected;
       If N<>nil then begin
          If Node=nil then TreeFormChange(nil, N)
                      else If N.AbsoluteIndex<>Node.AbsoluteIndex then TreeFormChange(nil, N);
       end;

    finally
       {������������� ����}
       AutoScrollTree(@TreeForm);

       {����������� Action}
       If Sender<>nil then EnablAction;

       {���������� ����������� �����}
       TreeForm.Items.EndUpdate;
       TreeForm.OnChange:=TreeFormChange;
    end;
end;


{==============================================================================}
{========================    �������: OnKeyDown    ============================}
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
{=========================  �������: ON_DblClick  =============================}
{==============================================================================}
procedure TFMAIN.TreeFormDblClick(Sender: TObject);
var N : TTreeNode;
    P : TPoint;
begin
    {�������������}
    If TreeForm.Selected=nil then Exit;

    {������� ������ ��� ���������� ������� � ������� ������}
    If Sender<>nil then begin
       If GetCursorPos(P) = false then Exit;
       P := TreeForm.ScreenToClient(P);
       N := TreeForm.GetNodeAt(P.X, P.Y);
    end else begin
       If TreeForm.SelectionCount=0 then Exit;
       N:=TreeForm.Selected;
    end;
    If N=nil then Exit;

    {��������/�������������� �����}
    AOpenExecute(Sender);
end;


{==============================================================================}
{===================    �������: OnCustomDrawItem    ==========================}
{==============================================================================}
procedure TFMAIN.TreeFormCustomDrawItem(Sender: TCustomTreeView;
                 Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
    If Node.Selected then begin
       If WIN_VER < wvVista then begin
          {��� WinXP}
          Sender.Canvas.Brush.Color := clGreen;
          Sender.Canvas.Font.Color  := clWhite;
       end else begin
          {��� WinVista � Win7}
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
    {����������� Action}
   // EnablAction;
end;

procedure TFMAIN.TreeFormCollapsing(Sender: TObject; Node: TTreeNode; var AllowCollapse: Boolean);
begin
    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{===================   ��������� ���� ��������� � TREEFORM   ==================}
{==============================================================================}
function TFMAIN.KeySelForm: String;
var NForm : TTreeNode;
begin
    {�������������}
    Result := '';
    NForm  := TreeForm.Selected;
    If NForm=nil then Exit;
    If NForm.ImageIndex <> ICO_FORM_DETAIL_COL then Exit;
    Result := IntToStr(Integer(NForm.Parent.Parent.Data))+CH_SPR+
              IntToStr(Integer(NForm.Data))+CH_SPR+
              IntToStr(Integer(NForm.Parent.Data));
end;


{==============================================================================}
{======================   ������� ��������� SELFORM   =========================}
{==============================================================================}
procedure TFMAIN.ClearSelForm;
begin
    With SelForm do begin
       Node := nil;
    end;
end;


{==============================================================================}
{=====================   ���������� ��������� SELFORM   =======================}
{==============================================================================}
function TFMAIN.SetSelForm: Boolean;
begin
    {�������������}
    Result := false;
    ClearSelForm;

    {���������� ����}
    SelForm.Node := TreeForm.Selected;
    If SelForm.Node=nil then Exit;

    // SKey

    {������������ ���������}
    Result:=true;
end;


{==============================================================================}
{======================   ����������� ������ ������   =========================}
{==============================================================================}
procedure TFMAIN.CBFindChange(Sender: TObject);
begin
    EnablAction;
end;


{==============================================================================}
{======================   ����� ����������� ��������   ========================}
{==============================================================================}
procedure TFMAIN.AFindPrevExecute(Sender: TObject);
begin
    Finder(false);
end;

{==============================================================================}
{=======================   ����� ���������� ��������   ========================}
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
{==============   �������: ��������� ������ ��������������   ==================}
{==============================================================================}
procedure TFMAIN.TreeFormMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
    {�������������}
    If Button <> mbLeft        then Exit;
    If TreeForm.Selected = nil then Exit;

    {������������� ������ ������ ��� �������}
    Case TreeForm.Selected.ImageIndex of
    ICO_FORM_DETAIL_ROW, ICO_FORM_DETAIL_COL: TreeForm.BeginDrag(true, 0);
    end;
end;



