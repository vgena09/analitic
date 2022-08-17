{==============================================================================}
{=====================   —Œ¡€“»≈: »«Ã≈Õ≈Õ»≈ ‘Œ–Ã”À€   =========================}
{==============================================================================}
procedure TFPARAM.EFormulaChange(Sender: TObject);
begin
    If TView.Selected = nil then Exit;
    LParam[TView.Selected.Index].SFormula := EFormula.Text;
end;


{==============================================================================}
{=======================   —Œ¡€“»≈: ON_DBL_CLICK   ============================}
{==============================================================================}
procedure TFPARAM.EFormulaDblClick(Sender: TObject);
var ISel, I, I1, I2 : Integer;
    P               : TPoint;
    S, SText        : String;
begin
    {»ÌËˆË‡ÎËÁ‡ˆËˇ}
    If GetCursorPos(P) = false then Exit;
    P    := EFormula.ScreenToClient(P);
    ISel := LoWord(EFormula.Perform(EM_CHARFROMPOS, 0, MakeLParam(P.X, P.Y)));
    S    := EFormula.Text;
    If S='' then Exit;
    I1   := ISel;
    I2   := ISel;
    For I := ISel downto 1 do begin
       If S[I] = '[' then begin
          I1 := I-1;
          Break;
       end;
    end;
    For I := ISel to Length(S) do begin
       If S[I] = ']' then begin
          I2 := I;
          Break;
       end;
    end;
    {”ÒÚ‡Ì‡‚ÎË‚‡ÂÏ ‚˚‰ÂÎÂÌËÂ}
    EFormula.SelStart  := I1;
    EFormula.SelLength := I2 - I1;
end;


{==============================================================================}
{====================   —Œ¡€“»≈: Ã€ÿ‹ ƒ¬»∆≈“—ﬂ ¬ œŒÀ≈   =======================}
{==============================================================================}
procedure TFPARAM.EFormulaMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var ISel : Integer;
    P    : TPoint;

   function GetSelKod(const Str: String; const ISel: Integer): String;
   var I, I1, I2 : Integer;
   begin
       Result := '';
       If Str = '' then Exit;
       I1   := 1;
       I2   := Length(Str);
       For I := ISel downto 1 do begin
          If Str[I] = '[' then begin
             I1 := I + 1;
             Break;
          end;
       end;
       For I := ISel to Length(Str) do begin
          If Str[I] = ']' then begin
             I2 := I - 1;
             Break;
          end;
       end;
       Result := Copy(Str, I1, I2 - I1 + 1);
       If (Pos('[', Result) > 0) or (Pos(']', Result) > 0) then Result := '';
   end;

begin
    If GetCursorPos(P) = false then Exit;
    P    := EFormula.ScreenToClient(P);
    ISel := LoWord(EFormula.Perform(EM_CHARFROMPOS, 0, MakeLParam(P.X, P.Y)));

    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ' + FFMAIN.KeyToCaption(GetSelKod(EFormula.Text, ISel), CH_SPR, ' / ', false);
end;

{==============================================================================}
{======================   —Œ¡€“»≈: Ã€ÿ‹ œŒ »ƒ¿≈“ œŒÀ≈   =======================}
{==============================================================================}
procedure TFPARAM.EFormulaMouseLeave(Sender: TObject);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ';
end;


{==============================================================================}
{=================   —Œ¡€“»≈: »«Ã≈Õ≈Õ»≈ ”—ÀŒ¬»ﬂ –¿— –¿— »   ===================}
{==============================================================================}
procedure TFPARAM.CBColorChange(Sender: TObject);
begin
    If TView.Selected = nil then Exit;
    LParam[TView.Selected.Index].SColor := IntToStr(CBColor.ItemIndex);
end;


{==============================================================================}
{=========================   DRAG: ƒŒœ”—“»ÃŒ—“‹   =============================}
{==============================================================================}
procedure TFPARAM.EFormulaDragOver(Sender, Source: TObject; X, Y: Integer; State: TDragState; var Accept: Boolean);
begin
    Accept := false;
    If Source = FFMAIN.TreeForm then begin
       If FFMAIN.TreeForm.Selected = nil then Exit;
       Case FFMAIN.TreeForm.Selected.ImageIndex of
       ICO_FORM_DETAIL_COL: Accept := true;
       end;
     end;
end;


{==============================================================================}
{=========================   DRAG: ¬—“¿¬»“‹  Œƒ    ============================}
{==============================================================================}
procedure TFPARAM.EFormulaDragDrop(Sender, Source: TObject; X, Y: Integer);
var N    : TTreeNode;
    SKey : String;
begin
    {»ÌËˆË‡ÎËÁ‡ˆËˇ}
    N := FFMAIN.TreeForm.Selected;
    If N = nil then Exit;
    SKey := '[' + IntToStr(Integer(N.Parent.Parent.Data))+CH_SPR+
                  IntToStr(Integer(N.Data))+CH_SPR+
                  IntToStr(Integer(N.Parent.Data)) + ']';
    {«‡ÔËÒ˚‚‡ÂÏ ÁÌ‡˜ÂÌËÂ}
    EFormula.SelStart := LoWord(EFormula.Perform(EM_CHARFROMPOS, 0, MakeLParam(X, Y)));
    EFormula.SelText  := SKey;
end;


