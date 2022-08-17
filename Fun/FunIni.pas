unit FunIni;

interface
uses
   Winapi.Windows {ReadSectionsEx},
   System.Classes, System.SysUtils, System.Variants, System.IniFiles,
   Vcl.Dialogs, Vcl.Controls, Vcl.Forms, Vcl.Graphics,   
   IdGlobal,
   FunType;

type TFormSaveParam = set of (fspPosition, fspState);

procedure WriteGlobalBool(const Section, Key: String; const Val: Boolean);
procedure WriteGlobalString(const Section, Key: String; const Val: String);
procedure WriteGlobalInteger(const Section, Key: String; const Val: Integer);
function  ReadGlobalBool(const Section, Key: String; const Val0: Boolean): Boolean;
function  ReadGlobalString(const Section, Key: String; const Val0: String): String;
function  ReadGlobalInteger(const Section, Key: String; const Val0: Integer): Integer;

procedure WriteLocalBool(const Section, Key: String; const Val: Boolean);
procedure WriteLocalString(const Section, Key: String; const Val: String);
procedure WriteLocalInteger(const Section, Key: String; const Val: Integer);
function  ReadLocalBool(const Section, Key: String; const Val0: Boolean): Boolean;
function  ReadLocalString(const Section, Key: String; const Val0: String): String;
function  ReadLocalInteger(const Section, Key: String; const Val0: Integer): Integer;

function  ReadSectionIni(const PList: PStringList; const IniFileName: String; const Section: String): Integer;
function  ReadLSectionsIni(const PList: PStringList; const IniFileName: String): Integer;

function  WriteSListIni(const P: PStrings; const SPath, Section: String): Boolean;
function  ReadSListIni(const P: PStrings; const SPath, Section: String): Boolean;

function  WriteCBListIni(const P: PComboBox; const Section: String): Boolean;
function  ReadCBListIni(const P: PComboBox; const Section: String): Boolean;

procedure SaveFormIni(F: TForm; const FSP: TFormSaveParam);
procedure LoadFormIni(F: TForm; const FSP: TFormSaveParam);

var BlockSaveForm : Boolean = false;

implementation

uses FunConst, FunSys, FunFiles, FunText;


{==============================================================================}
{=====================  —Œ’–¿Õ≈Õ»≈ √ÀŒ¡¿À‹Õ€’ œ¿–¿Ã≈“–Œ¬  =====================}
{==============================================================================}
procedure WriteGlobalBool(const Section, Key: String; const Val: Boolean);
var FIni : TIniFile;   
begin
    FIni := TIniFile.Create(FPATH_PROG_INI);
    try     FIni.WriteBool(Section, Key, Val);
    finally FIni.Free; end;
end;

procedure WriteGlobalString(const Section, Key: String; const Val: String);
var FIni : TIniFile;
begin
    FIni := TIniFile.Create(FPATH_PROG_INI);
    try     FIni.WriteString(Section, Key, Val);
    finally FIni.Free; end;
end;

procedure WriteGlobalInteger(const Section, Key: String; const Val: Integer);
var FIni : TIniFile;
begin
    FIni := TIniFile.Create(FPATH_PROG_INI);
    try     FIni.WriteInteger(Section, Key, Val);
    finally FIni.Free; end;
end;

{==============================================================================}
{=======================  ◊“≈Õ»≈ √ÀŒ¡¿À‹Õ€’ œ¿–¿Ã≈“–Œ¬  =======================}
{==============================================================================}
function ReadGlobalBool(const Section, Key: String; const Val0: Boolean): Boolean;
var FIni : TIniFile;
begin
    Result := Val0;
    FIni := TIniFile.Create(FPATH_PROG_INI);            
    try     Result:=FIni.ReadBool(Section, Key, Val0);
    finally FIni.Free; end;
end;

function ReadGlobalString(const Section, Key: String; const Val0: String): String;
var FIni : TIniFile;
begin
    Result := Val0;
    FIni := TIniFile.Create(FPATH_PROG_INI);
    try     Result:=FIni.ReadString(Section, Key, Val0);
    finally FIni.Free; end;
end;

function ReadGlobalInteger(const Section, Key: String; const Val0: Integer): Integer;
var FIni : TIniFile;
begin
    Result := Val0;
    FIni := TIniFile.Create(FPATH_PROG_INI);
    try     Result:=FIni.ReadInteger(Section, Key, Val0);
    finally FIni.Free; end;
end;


{==============================================================================}
{======================  —Œ’–¿Õ≈Õ»≈ ÀŒ ¿À‹Õ€’ œ¿–¿Ã≈“–Œ¬  =====================}
{==============================================================================}
procedure WriteLocalBool(const Section, Key: String; const Val: Boolean);
var FIni : TIniFile;
begin
    FIni := TIniFile.Create(PATH_WORK_INI);
    try     FIni.WriteBool(Section, Key, Val);
    finally FIni.Free; end;
end;

procedure WriteLocalString(const Section, Key: String; const Val: String);
var FIni : TIniFile;
begin
    FIni := TIniFile.Create(PATH_WORK_INI);
    try     FIni.WriteString(Section, Key, Val);
    finally FIni.Free; end;
end;

procedure WriteLocalInteger(const Section, Key: String; const Val: Integer);
var FIni : TIniFile;
begin
    FIni := TIniFile.Create(PATH_WORK_INI);
    try     FIni.WriteInteger(Section, Key, Val);
    finally FIni.Free; end;
end;


{==============================================================================}
{========================  ◊“≈Õ»≈ ÀŒ ¿À‹Õ€’ œ¿–¿Ã≈“–Œ¬  =======================}
{==============================================================================}
function ReadLocalBool(const Section, Key: String; const Val0: Boolean): Boolean;
var FIni : TIniFile;
begin
    Result := Val0;
    FIni := TIniFile.Create(PATH_WORK_INI);
    try     Result:=FIni.ReadBool(Section, Key, Val0);
    finally FIni.Free; end;
end;

function ReadLocalString(const Section, Key: String; const Val0: String): String;
var FIni : TIniFile;
begin
    Result := Val0;
    FIni := TIniFile.Create(PATH_WORK_INI);
    try     Result:=FIni.ReadString(Section, Key, Val0);
    finally FIni.Free; end;
end;

function ReadLocalInteger(const Section, Key: String; const Val0: Integer): Integer;
var FIni : TIniFile;
begin
    Result := Val0;
    FIni := TIniFile.Create(PATH_WORK_INI);
    try     Result:=FIni.ReadInteger(Section, Key, Val0);
    finally FIni.Free; end;
end;


{==============================================================================}
{=========================  ◊“≈Õ»≈ —Œƒ≈–∆¿Õ»ﬂ —≈ ÷»»  =========================}
{==============================================================================}
function ReadSectionIni(const PList: PStringList; const IniFileName: String; const Section: String): Integer;
var FIni : TIniFile;
begin
    Result := -1;
    FIni := TIniFile.Create(IniFileName);
    try     PList^.BeginUpdate;
            FIni.ReadSection(Section, PList^);
            Result:=PList^.Count;
    finally FIni.Free;
            PList^.EndUpdate;
    end;
end;


{==============================================================================}
{==========================  ◊“≈Õ»≈ —œ»— ¿ —≈ ÷»…  ============================}
{==============================================================================}
function ReadLSectionsIni(const PList: PStringList; const IniFileName: String): Integer;
var FIni : TIniFile;
begin
    Result := -1;
    FIni := TIniFile.Create(IniFileName);
    try     PList^.BeginUpdate;
            FIni.ReadSections(PList^);
            Result:=PList^.Count;
    finally FIni.Free;
            PList^.EndUpdate;
    end;
end;

{==============================================================================}
{====================    —Œ’–¿Õ≈Õ»≈ —œ»— ¿ TSTRINGS    ========================}
{==============================================================================}
{==================== ≈ÒÎË SPath = '' - ÎÓÍ‡Î¸Ì˚È Ù‡ÈÎ ========================}
{==============================================================================}
function WriteSListIni(const P: PStrings; const SPath, Section: String): Boolean;
var FIni : TIniFile;
    S    : String;
    I    : Integer;
begin
    {»ÌËˆË‡ÎËÁ‡ˆËˇ}
    Result := false;
    If P = nil then Exit;
    If SPath = '' then S := PATH_WORK_INI
                  else S := SPath;
    FIni := TIniFile.Create(S);
    try
       P^.BeginUpdate;
       FIni.EraseSection(Section);
       For I:=1 to P^.Count do FIni.WriteString(Section, IntToStr(I), P^[I-1]);
       Result:=true;
    finally
       FIni.Free;
       P^.EndUpdate;
    end;
end;


{==============================================================================}
{====================     «¿√–”« ¿ —œ»— ¿ TSTRINGS     ========================}
{==============================================================================}
{==================== ≈ÒÎË SPath = '' - ÎÓÍ‡Î¸Ì˚È Ù‡ÈÎ ========================}
{==============================================================================}
function ReadSListIni(const P: PStrings; const SPath, Section: String): Boolean;
var FIni : TIniFile;
    S    : String;
    I    : Integer;
begin
    {»ÌËˆË‡ÎËÁ‡ˆËˇ}
    Result := false;
    If P = nil then Exit;
    If SPath = '' then S := PATH_WORK_INI
                  else S := SPath;
    P^.BeginUpdate;
    try
       P^.Clear;
       {◊ËÚ‡ÂÏ ÒÂÍˆË˛}
       FIni := TIniFile.Create(S);
       try     FIni.ReadSectionValues(Section, P^);
       finally FIni.Free;
       end;
       {”·Ë‡ÂÏ ÍÎ˛˜Ë}
       For I:=0 to P^.Count-1 do begin
          S := P^[I];
          TokChar(S, '=');
          If S<>'' then P^[I] := S;
       end;
       Result := true;
    finally
       P^.EndUpdate;
    end;
end;


{==============================================================================}
{======================   —Œ’–¿Õ≈Õ»≈ —œ»— ¿ TCOMBOBOX   =======================}
{==============================================================================}
function WriteCBListIni(const P: PComboBox; const Section: String): Boolean;
begin
    Result := WriteSListIni(@P^.Items, '', Section);
end;


{==============================================================================}
{========================   «¿√–”« ¿ —œ»— ¿ TCOMBOBOX   =======================}
{==============================================================================}
function ReadCBListIni(const P: PComboBox; const Section: String): Boolean;
begin
    {»ÌËˆË‡ÎËÁ‡ˆËˇ}
    Result := false;
    If P=nil then Exit;
    P^.Items.BeginUpdate;
    try
       {◊ËÚ‡ÂÏ ÚÂÍÒÚ}
       If Not ReadSListIni(@P^.Items, '', Section) then Exit;
       {¬ÓÒÒÚ‡Ì‡‚ÎË‚‡ÂÏ ÚÂÍÒÚ}
       If P^.Items.Count > 0 then P^.Text := P^.Items[0];
       Result := true;
    finally
       P^.Items.EndUpdate;
    end;
end;


{==============================================================================}
{================  —Œ’–¿Õ≈Õ»≈ œ¿–¿Ã≈“–Œ¬ ‘Œ–Ã€ ¬ INI-‘¿…À   ===================}
{==============================================================================}
procedure SaveFormIni(F: TForm; const FSP: TFormSaveParam);
var Section: String;     
begin
    If BlockSaveForm then Exit;
    Section := INI_FORM_PARAM+F.Name;
    With F do begin
       If fspPosition in FSP then begin
          WriteLocalInteger(Section, INI_FORM_PARAM_LEFT,   Left);
          WriteLocalInteger(Section, INI_FORM_PARAM_TOP,    Top);
          WriteLocalInteger(Section, INI_FORM_PARAM_WIDTH,  Width);
          WriteLocalInteger(Section, INI_FORM_PARAM_HEIGHT, Height);
       end;
       If fspState in FSP then begin
          WriteLocalBool(Section, INI_FORM_PARAM_MAXIMIZE, WindowState = wsMaximized);
       end;
    end;
end;


{==============================================================================}
{================  «¿√–”« ¿ œ¿–¿Ã≈“–Œ¬ ‘Œ–Ã€ »« INI-‘¿…À¿   ===================}
{==============================================================================}
procedure LoadFormIni(F: TForm; const FSP: TFormSaveParam);
var Section: String;     
begin
    Section := INI_FORM_PARAM+F.Name;
    BlockSaveForm := true;
    try
       With F do begin
          If fspPosition in FSP then begin
             Left   := ReadLocalInteger(Section, INI_FORM_PARAM_LEFT,   Left);
             Top    := ReadLocalInteger(Section, INI_FORM_PARAM_TOP,    Top);
             Width  := ReadLocalInteger(Section, INI_FORM_PARAM_WIDTH,  Width);
             Height := ReadLocalInteger(Section, INI_FORM_PARAM_HEIGHT, Height);
          end;
          If fspState in FSP then begin
             If ReadLocalBool(Section, INI_FORM_PARAM_MAXIMIZE, true) then WindowState := wsMaximized
                                                                      else WindowState := wsNormal;
          end;
       end;
    finally
       BlockSaveForm := false;
    end;
end;

end.

