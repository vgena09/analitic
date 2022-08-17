{==============================================================================}
{==========================   «¿œ»—‹ œ¿–¿Ã≈“–Œ¬   =============================}
{==============================================================================}
function TFPARAM.ParamWrite(const SPath: String): Boolean;
var SList : TStringList;
    S     : String;
    Ind   : Integer;
begin
    Result := false;
    SList  := TStringList.Create;
    try
       SList.Clear;
       For Ind := Low(LParam) to High(LParam) do begin
          S := LParam[Ind].SCaption + SPR_SAV_STR + LParam[Ind].SFormula + SPR_SAV_STR + LParam[Ind].SColor;
          SList.Add(S);
       end;
       Result := WriteSListIni(@SList, SPath, INI_ANALITIC_PARAM);
    finally
       SList.Free;
    end;
end;


{==============================================================================}
{==========================   ◊“≈Õ»≈ œ¿–¿Ã≈“–Œ¬   =============================}
{==============================================================================}
function TFPARAM.ParamRead(const SPath: String): Boolean;
var SList : TStringList;
    S     : String;
    Ind   : Integer;
begin
    ParamClear;
    Result := false;
    SList  := TStringList.Create;
    try
       ReadSListIni(@SList, SPath, INI_ANALITIC_PARAM);
       SetLength(LParam, SList.Count);
       For Ind := Low(LParam) to High(LParam) do begin
          S := SList[Ind];
          LParam[Ind].SCaption := TokStr(S, SPR_SAV_STR);
          LParam[Ind].SFormula := TokStr(S, SPR_SAV_STR); If S='' then S:='0';
          LParam[Ind].SColor   := S;
       end;
       Result := true;
    finally
       SList.Free;
       TViewRefresh(-1);
       EnablAction;
    end;
end;


{==============================================================================}
{==========================   Œ◊»—“ ¿ œ¿–¿Ã≈“–Œ¬   ============================}
{==============================================================================}
procedure TFPARAM.ParamClear;
begin
    SetLength(LParam, 0);
    TViewRefresh(-1);
    EnablAction;
end;







