{==============================================================================}
{=======================    СОБЫТИЕ: ON_DBL_CLICK    ==========================}
{==============================================================================}
procedure TFPARAM.TViewDblClick(Sender: TObject);
begin
    TView.Selected.EditText;
end;


{==============================================================================}
{========================    СОБЫТИЕ: ON_CHANGE    ============================}
{==============================================================================}
procedure TFPARAM.TViewChange(Sender: TObject; Node: TTreeNode);
var SFormula, SColor : String;
    Enabl            : Boolean;
begin
    {Инициализация}
    SFormula := '';
    SColor   := '0';
    Enabl    := false;
    try
       If Node = nil then Exit;
       If Not Node.Selected then Exit;
       SFormula := LParam[Node.Index].SFormula;
       SColor   := LParam[Node.Index].SColor;
       Enabl    := true;
    finally
       EFormula.Text     := SFormula;
       EFormula.Enabled  := Enabl;
       AFormula.Enabled  := Enabl;
       LFormula.Enabled  := Enabl;
       If SColor = '' then SColor := '0';
       CBColor.ItemIndex := StrToInt(SColor);
       CBColor.Enabled   := Enabl;
       LColor.Enabled    := Enabl;
       // AutoScrollTree(@TView);
       // Сдвиг вправо       TView.Perform(WM_HSCROLL, SB_LINELEFT, 10);
    end;
end;


{==============================================================================}
{==============    СОБЫТИЕ: ЗАЕРШЕНИЕ РЕДАКТИРОВАНИЯ ITEM    ==================}
{==============================================================================}
procedure TFPARAM.TViewEdited(Sender: TObject; Node: TTreeNode; var S: string);
begin
    If Node = nil then Exit;
    LParam[Node.Index].SCaption := Trim(S);
end;


{==============================================================================}
{=========================   DRAG: ДОПУСТИМОСТЬ   =============================}
{==============================================================================}
procedure TFPARAM.TViewDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
var N: TTreeNode;
begin
    {Скроллинг}
    If (Y > 0) and (Y < 30) then SendMessage(TView.Handle, WM_VSCROLL, SB_LINEUP, 0);
    If (Y > TView.ClientHeight - 30) and (Y < TView.ClientHeight) then SendMessage(TView.Handle, WM_VSCROLL, SB_LINEDOWN, 0);
    If (X > 0) and (X < 30) then SendMessage(TView.Handle, WM_HSCROLL, SB_LINELEFT, 0);
    If (X > TView.ClientHeight - 30) and (X < TView.ClientWidth) then SendMessage(TView.Handle, WM_HSCROLL, SB_LINERIGHT, 0);

    {Перемещение параметра внутри компонента}
    If Source = TView then begin
       Accept := false;
       If TView.Items.Count < 2 then Exit;
       N := TView.GetNodeAt(X, Y);
       If N = nil then N := TView.Items[High(LParam)];
       Accept := N.Index <> TView.Selected.Index;
    end;

    {Импорт параметра - только таблицы, строки или столбцы}
    If Source = FFMAIN.TreeForm then begin
       Accept := false;
       If FFMAIN.TreeForm.Selected = nil then Exit;
       Case FFMAIN.TreeForm.Selected.ImageIndex of
       ICO_FORM_DETAIL_ROW, ICO_FORM_DETAIL_COL: Accept := true;
       end;
     end;
end;


{==============================================================================}
{====================       DRAG: ПЕРЕМЕЩАЕМ ITEM        ======================}
{==============================================================================}
procedure TFPARAM.TViewDragDrop(Sender, Source: TObject; X, Y: Integer);
var IndOld, IndNew, Ind : Integer;
    ParamOld            : TParam;
    N, N0               : TTreeNode;

    {Добавляем Item}
    procedure AddItem(const SFormula: String);
    var Ind: Integer;
    begin
        If Length(LParam) >= ITEM_MAX then Exit;
        SetLength(LParam, Length(LParam) + 1);
        If IndNew = -1 then IndNew := High(LParam) else Inc(IndNew);

        For Ind := High(LParam) downto IndNew + 1 do LParam[Ind] := LParam[Ind - 1];
        LParam[IndNew].SFormula := '[' + SFormula + ']';
        LParam[IndNew].SCaption := FFMAIN.FormulaToCaption(SFormula, true);
        LParam[IndNew].SColor   := '0'
    end;
begin
    {Инициализация}
    N := TView.GetNodeAt(X, Y);
    If (N = nil) and  (Length(LParam) > 0) then N := TView.Items[High(LParam)];
    If N <> nil then IndNew := N.Index else IndNew := -1;

    {Перемещение параметра внутри компонента}
    If Source = TView then begin
       If TView.Selected    = nil then Exit;
       If TView.Items.Count < 2   then Exit;
       IndOld   := TView.Selected.Index;
       ParamOld := LParam[IndOld];
       For Ind  := IndOld       to     High(LParam) - 1 do LParam[Ind] := LParam[Ind + 1];
       For Ind  := High(LParam) downto IndNew + 1       do LParam[Ind] := LParam[Ind - 1];
       LParam[IndNew] := ParamOld;
    end;

    {Импорт параметра}
    If Source = FFMAIN.TreeForm then begin
       {Допустимость источника}
       N := FFMAIN.TreeForm.Selected;
       If N = nil then Exit;
       {Добавляем Item(s)}
       Case FFMAIN.TreeForm.Selected.ImageIndex of
       ICO_FORM_DETAIL_ROW: begin
             N0 := N.GetFirstChild;
             While N0 <> nil do begin
                AddItem(IntToStr(Integer(N.Parent.Data))+CH_SPR+
                        IntToStr(Integer(N0.Data))+CH_SPR+
                        IntToStr(Integer(N.Data)));
                N0 := N.GetNextChild(N0);
             end;
          end;
       ICO_FORM_DETAIL_COL: begin
             AddItem(IntToStr(Integer(N.Parent.Parent.Data))+CH_SPR+
                     IntToStr(Integer(N.Data))+CH_SPR+
                     IntToStr(Integer(N.Parent.Data)));
          end;
       end;
    end;

    {Обновляем список}
    TViewRefresh(IndNew);
end;


{==============================================================================}
{================   DRAG: УДАЛЯЕМ ITEM ПРИ ЕГО ВЫТАСКИВАНИИ    ================}
{==============================================================================}
procedure TFPARAM.TViewEndDrag(Sender, Target: TObject; X, Y: Integer);
var P : TPoint;
begin
    If Target <> nil   then Exit;
    If Sender <> TView then Exit;
    If GetCursorPos(P) = false then Exit;
    P := TView.ScreenToClient(P);
    If (P.X >= 0) and (P.X <= TView.ClientWidth) and (P.Y >= 0) and (P.Y <= TView.ClientHeight) then Exit;
    ADelExecute(Sender);
end;


{==============================================================================}
{===================    СОБЫТИЕ: ON_CUSTOM_DRAW_ITEM    =======================}
{==============================================================================}
procedure TFPARAM.TViewCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
begin
    If Node.Selected then begin
       If FFMAIN.WIN_VER < wvVista then begin
          {Для WinXP}
          Sender.Canvas.Brush.Color := clGreen;
          Sender.Canvas.Font.Color  := clWhite;
       end else begin
          {Для WinVista и Win7}
          Sender.Canvas.Font.Color := clRed;
       end;
    end else begin
       Sender.Canvas.Font.Color := clTeal;
    end;
end;


{==============================================================================}
{=============   ОБНОВЛЕНИЕ ОТОБРАЖЕНИЯ СПИСКА ПАРАМЕТРОВ   ===================}
{==============================================================================}
procedure TFPARAM.TViewRefresh(const ISel: Integer);
var N   : TTreeNode;
    Ind : Integer;
begin
    TView.Items.BeginUpdate;
    try
       TView.OnChange := nil;
       TView.Items.Clear;
       For Ind := Low(LParam) to High(LParam) do begin
          N := TView.Items.AddChildObject(nil, LParam[Ind].SCaption, Pointer(Ind)); // Data не используется
          N.ImageIndex    := ICO_SYS_FORMULA;
          N.SelectedIndex := ICO_SYS_FORMULA;
          If Ind = ISel then TView.Selected := N;
       end;
    finally
        TView.OnChange := TViewChange;
        TView.Items.EndUpdate;
        TViewChange(nil, TView.Selected);
        EnablAction;
    end;
end;


