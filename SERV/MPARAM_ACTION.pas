{==============================================================================}
{=========================   ƒŒ—“”œÕŒ—“‹ ACTION   =============================}
{==============================================================================}
procedure TFPARAM.EnablAction;
var IsClear, IsAdd, IsDel, IsSave : Boolean;
begin
    FFMAIN.EnablAction;
    IsClear := Length(LParam) > 0;
    IsAdd   := Length(LParam) < ITEM_MAX;
    IsDel   := TView.Selected <> nil;
    IsSave  := Length(LParam) > 0;

    AClear.Enabled := IsClear;
    AAdd.Enabled   := IsAdd;
    ADel.Enabled   := IsDel;
    ASave.Enabled  := IsSave;
end;


{==============================================================================}
{======================   ACTION: —Œ’–¿Õ»“‹ œ¿–¿Ã≈“–€  ========================}
{==============================================================================}
procedure TFPARAM.ASaveExecute(Sender: TObject);
var FName: String;
begin
    SaveDlg.InitialDir := PATH_WORK;
    If Not SaveDlg.Execute then Exit;
    FName := SaveDlg.FileName;
    ParamWrite(FName);
end;


{==============================================================================}
{====================   ACTION: ¬Œ——“¿ÕŒ¬»“‹ œ¿–¿Ã≈“–€  =======================}
{==============================================================================}
procedure TFPARAM.AOpenAsExecute(Sender: TObject);
var FName: String;
begin
    OpenDlg.InitialDir := PATH_WORK;
    OpenDlg.FileName   := '';
    If Not OpenDlg.Execute then Exit;
    FName := OpenDlg.FileName;
    ParamRead(FName);
end;


procedure TFPARAM.AOpenExecute(Sender: TObject);
var FName: String;
begin
    OpenDlg.InitialDir := PATH_BD_SHABLON;
    OpenDlg.FileName   := '';
    If Not OpenDlg.Execute then Exit;
    FName := OpenDlg.FileName;
    ParamRead(FName);
end;


{==============================================================================}
{======================   ACTION: Œ◊»—“»“‹ œ¿–¿Ã≈“–€  =========================}
{==============================================================================}
procedure TFPARAM.AClearExecute(Sender: TObject);
begin
    If Length(LParam)=0 then Exit;
    If MessageDlg('œÓ‰Ú‚Â‰ËÚÂ Ó˜ËÒÚÍÛ ÒÔËÒÍ‡ Ô‡‡ÏÂÚÓ‚!', mtWarning, [mbYes, mbNo], 0) <> mrYes then Exit;
    ParamClear;
end;


{==============================================================================}
{=====================   ACTION: ƒŒ¡¿¬»“‹ œ¿–¿Ã≈“–   ==========================}
{==============================================================================}
procedure TFPARAM.AAddExecute(Sender: TObject);
var Ind: Integer;
begin
    Ind := Length(LParam);
    SetLength(LParam, Ind + 1);
    With LParam[Ind] do begin
       SCaption := '';
       SFormula := '';
       SColor   := '0';
    end;
    TViewRefresh(Ind);
end;


{==============================================================================}
{=======================  ACTION: ”ƒ¿À»“‹ œ¿–¿Ã≈“–   ==========================}
{==============================================================================}
procedure TFPARAM.ADelExecute(Sender: TObject);
var Ind, IndDel : Integer;
begin
    If TView.Selected = nil then Exit;
    IndDel := TView.Selected.Index;
    For Ind  := IndDel to High(LParam) - 1 do LParam[Ind] := LParam[Ind + 1];
    SetLength(LParam, Length(LParam) - 1);
    If IndDel > High(LParam) then IndDel := High(LParam);
    TViewRefresh(IndDel);
end;


{==============================================================================}
{====================   ACTION: –≈ƒ¿ “»–Œ¬¿“‹ ‘Œ–Ã”À”   =======================}
{==============================================================================}
procedure TFPARAM.AFormulaExecute(Sender: TObject);
var F            : TFKOD;
    S, SSel      : String;
    IStart, ILen : Integer;
begin
    IStart := EFormula.SelStart;
    ILen   := EFormula.SelLength;
    S      := CutModulChar(EFormula.SelText, '[', ']');
    If S='' then SSel := EFormula.SelText
            else SSel := S;
    F     := TFKOD.Create(Self);
    try S := F.Execute(SSel, false);
        EFormula.SetFocus;
        EFormula.SelStart  := IStart;
        EFormula.SelLength := ILen;
        If S<>SSel then EFormula.SelText := '['+S+']';
    finally F.Free;
    end;
end;

