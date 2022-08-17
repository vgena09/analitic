{==============================================================================}
{=========            Вырезает первый пакет SPack из списка             =======}
{==============================================================================}
(*========          {  {  SPACKTYPE  {  SPACKTYPE  {  } } } }           ======*)
{=========              ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^              =======}
{=========                              SPack                           =======}
{=========              SPackType = SPackTypeList[Result]               =======}
{==============================================================================}
function CutModulPackList(const Str           : String;
                          const SPackTypeList : array of String;
                          var   SPack         : String): Integer;
var I, Ind, IndMin: Integer;
begin
    {Инициализация}
    Result := -1;
    IndMin := -1;
    SPack  := '';

    {Ищем первый из списка идентификатор пакета}
    For I:=Low(SPackTypeList) to High(SPackTypeList) do begin
       Ind:=GetPosStartPack(Str, SPackTypeList[I]);
       If Ind = -1    then Continue;
       If (Ind>=IndMin) and (IndMin>-1) then Continue;
       IndMin:=Ind;
       Result:=I;
    end;
    If Result=-1 then Exit;

    {Определяем конец пакета}
    I:=GetPosEndPack(Str, SPackTypeList[Result]);
    If I=-1 then begin Result:=-1; Exit; end;

    {Вырезаем пакет}
    SPack:=Copy(Str, IndMin+1, I-IndMin-1);
end;


{==============================================================================}
{=========             Вырезает первый пакет типа SPackType             =======}
{==============================================================================}
(*========          {  {  SPACKTYPE  {  SPACKTYPE  {  } } } }           ======*)
{=========              ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^              =======}
{=========                             Result                           =======}
{==============================================================================}
function CutModulPack(const Str, SPackType: String): String;
var I1, I2: Integer;
begin
    Result:='';
    I1:=GetPosStartPack(Str, SPackType); If I1=-1 then Exit;
    I2:=GetPosEndPack  (Str, SPackType); If I2=-1 then Exit;
    Result:=Copy(Str, I1+1, I2-I1-1);
end;


{==============================================================================}
{=====  Вырезает первый вложенный блок, ограниченный символами С1 и С2  =======}
{==============================================================================}
function CutModulChar(const S: String; const C1, C2: Char): String;
var I1, I2: Integer;
begin
    Result:='';
    I1:=GetPosStartChar(S, C1,  C2);   if I1=-1 then Exit;
    I2:=GetPosEndChar  (S, I1+1,C2);   if I2=-1 then Exit;
    Result:=Copy(S, I1+1, I2-(I1+1));
end;


{==============================================================================}
{=====  Вырезает первый вложенный блок, ограниченный строками S1 и S2  ========}
{==============================================================================}
function CutModulStr(const S: String; const S1, S2: String): String;
var I1, I2: Integer;
begin
    Result:='';
    If (S='') or (S1='') or (S2='') then Exit;
    I1:=GetPosStartStr(S, S1,   S2);   if I1=-1 then Exit;
    I2:=GetPosEndStr  (S, I1+1, S2);   if I2=-1 then Exit;
    Result:=Copy(S, I1, I2-I1+1);
end;


{==============================================================================}
{======  Заменяет первый вложенный блок, ограниченный символами С1 и С2  ======}
{==============================================================================}
function ReplModulChar(const Str1, Str2: String; const C1, C2: Char): String;
var I1, I2: Integer;
begin
    Result:=Str1;
    I1:=GetPosStartChar(Str1, C1, C2);   if I1=-1 then Exit;
    I2:=GetPosEndChar  (Str1, I1+1, C2); if I2=-1 then Exit;
    Delete(Result, I1, I2-I1+1);
    Insert(Str2, Result, I1);
end;


{==============================================================================}
{======  Заменяет первый вложенный блок, ограниченный строками S1 и S2  =======}
{==============================================================================}
function ReplModulStr(const Str1, Str2: String; const S1, S2: String): String;
var I1, I2: Integer;
begin
    Result:=Str1;
    If (Str1='') or (S1='') or (S2='') then Exit;
    I1:=GetPosStartStr(Str1, S1, S2)-Length(S1);   if I1=-1 then Exit;
    I2:=GetPosEndStr  (Str1, I1+1, S2)+Length(S2); if I2=-1 then Exit;
    Delete(Result, I1, I2-I1+1);
    Insert(Str2, Result, I1);
end;


{==============================================================================}
{================   Заменяет первый пакет типа SPackType   ====================}
{==============================================================================}
(*========          {  {  SPACKTYPE  {  SPACKTYPE  {  } } } }           ======*)
{=========              ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^              =======}
{==============================================================================}
function ReplModulPack(const Str1, Str2, SPackType: String): String;
var I1, I2: Integer;
begin
    Result:=Str1;
    I1:=GetPosStartPack(Str1, SPackType);   If I1=-1 then Exit;
    I2:=GetPosEndPack  (Str1, SPackType);   If I2=-1 then Exit;
    Delete(Result, I1, I2-I1+1);
    Insert(Str2, Result, I1);
end;


{==============================================================================}
{=======  Определяет позицию первого символа первого вложенного блока,  =======}
{=======                   ограниченного символами С1-С2                =======}
{==============================================================================}
function GetPosStartChar(const S: String; const C1, C2: Char): Integer;
var I, Num, MaxNum: Integer;
begin
    Result:=-1; Num:=0; MaxNum:=0;
    for I:=1 to Length(S) do begin
        if S[I]=C1 then begin
           Inc(Num);
           if Num>MaxNum then begin
              MaxNum:=Num;
              Result:=I;
           end;
        end;
        if S[I]=C2 then Dec(Num);
    end;
end;


{==============================================================================}
{=======  Определяет позицию первого символа первого вложенного блока,  =======}
{=======                   ограниченного строками S1-S2                 =======}
{==============================================================================}
function GetPosStartStr(const S: String; const S1, S2: String): Integer;
label Nx;
var I, I1, I2, Num, MaxNum: Integer;
begin
    Result:=-1; Num:=0; MaxNum:=0;
    If (S1='') or (S2='') or (S1=S2) then Exit;
    I:=1;

Nx: I1:=InStrMy(I, S, S1);
    I2:=InStrMy(I, S, S2);
    If I1<I2 then begin
       Inc(Num);
       if Num>MaxNum then begin
          MaxNum:=Num;
          Result:=I1+Length(S1);
       end;
       I:=I2+Length(S2);
    end;
    If I1>I2 then begin
       Dec(Num);
       I:=I2+Length(S2);
    end;
    If (I1 > 0) and (I2 > 0) then goto Nx;
end;


{==============================================================================}
{====    Определяет позицию первого символа первого пакета SPACKTYPE    =======}
{==============================================================================}
(*===  {  {  SPACKTYPE {  SPACKTYPE {  } } } }                          ======*)
{====     +                                                             =======}
{==============================================================================}
function GetPosStartPack(const S, SPackType: String): Integer;
label Nx;
var S0, SPackType0 : String;
    I0, I, IStart  : Integer;
begin
    {Инициализация}
    Result     := -1;
    IStart     := 1;
    S0         := AnsiUpperCase(S);
    SPackType0 := AnsiUpperCase(SPackType);

    {Определяем положение ключевой фразы}
Nx: I:=InStrMy(IStart, S0, SPackType0);
    If I<1 then Exit;
    IStart:=I+Length(SPackType0);

    (*Перед ключевой фразой обязательно д/б {  *)
    S0:=Copy(S0, 1, I-1);
    For I0:=Length(S0) downto 1 do begin
       If S0[I0]='{' then begin
          Result:=I0;
          Break;
       end;
       If S0[I0]<>' ' then Break;
    end;

    If Result=-1 then Goto Nx;
end;


{==============================================================================}
{=====  Определяет позицию последнего символа первого вложенного блока,  ======}
{=====                   ограниченного символами С1-С2                   ======}
{==============================================================================}
function GetPosEndChar(const S: String; const StartPos: Integer; const C2:Char): Integer;
var I: Integer;
begin
    Result:=-1;
    for I:=StartPos to Length(S) do begin
        if S[I]=C2 then begin Result:=I; Exit; end;
    end;
end;


{==============================================================================}
{=====  Определяет позицию последнего символа первого вложенного блока,  ======}
{=====                   ограниченного строками S1-S2                    ======}
{==============================================================================}
function GetPosEndStr(const S: String; const StartPos: Integer; const S2: String): Integer;
var I: Integer;
begin
    Result:=-1;
    If S2='' then Exit;

    I:=InStrMy(StartPos, S, S2);
    If I>0 then Result:=I-1;
end;


{==============================================================================}
{====    Определяет позицию последнего символа первого пакета SPackType     ===}
{==============================================================================}
(*===  {  {  SPACKTYPE {  SPACKTYPE {  } } } }                              ==*)
{====                                      +                                ===}
{==============================================================================}
function GetPosEndPack(const S, SPackType: String): Integer;
var S0            : String;
    I0, I, ICount : Integer;
begin
    {Инициализация}
    Result := -1;
    S0     := AnsiUpperCase(S);
    I0     := GetPosStartPack(S0, SPackType);
    If I0<1 then Exit;

    {Ищем пару к закрывающей скобке}
    ICount:=0;
    For I:=I0 to Length(S0) do begin
       If (S0[I]<>'{') and (S0[I]<>'}') then Continue;
       If S0[I]='{' then Inc(ICount) else Dec(ICount);
       If ICount=0 then begin
          Result:=I;
          Break;
       end;
    end;
end;


