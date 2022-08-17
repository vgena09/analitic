unit MKOD_CORRECT;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ActnList,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons, Vcl.ComCtrls,
  MAIN;

type
  TFKOD_CORRECT = class(TForm)
    ActionList1: TActionList;
    ARun: TAction;
    ACancel: TAction;
    POld: TPanel;
    Label1: TLabel;
    PNew: TPanel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    ETab2: TEdit;
    ECol2: TEdit;
    ERow2: TEdit;
    BtnOld: TButton;
    AKodFormula: TAction;
    BtnNewKod: TButton;
    AKodNew: TAction;
    PBottom: TPanel;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    Bevel2: TBevel;
    Bevel3: TBevel;
    Panel1: TPanel;
    Panel2: TPanel;
    AKodOld: TAction;
    Button1: TButton;
    Panel3: TPanel;
    CBFormula: TComboBox;
    procedure ARunExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EditChange(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure AKodFormulaExecute(Sender: TObject);
    procedure AKodNewExecute(Sender: TObject);
    procedure AKodOldExecute(Sender: TObject);
  private
    FFMAIN                                                  : TFMAIN;
    IsFormula, IsReplace, IsDelTab, IsDelTabEmpty, IsDelKey : Boolean;
    STab1, SRow1, SCol1, SFormula                           : String;
    STab2, SRow2, SCol2                                     : String;

    procedure LoopDir(const SPath: String);
    function  StepFile(const SFile: String): Boolean;
  public
    procedure Execute(const STab, SCol, SRow: String);
  end;

var
  FKOD_CORRECT: TFKOD_CORRECT;

implementation

uses FunConst, FunText, FunSys, FunIni,
     MKOD, MKOD_FORMULA;

{$R *.dfm}


{==============================================================================}
{==========================   СОЗДАНИЕ ФОРМЫ   ================================}
{==============================================================================}
procedure TFKOD_CORRECT.FormCreate(Sender: TObject);
begin
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' Перекодировка базы данных';
    
    {Восстанавливаем историю}
    CBFormula.DropDownCount := 25;
    ReadCBListIni(@CBFormula, INI_KOD_CORRECT);
end;


{==============================================================================}
{==========================   ЗАКРЫТИЕ ФОРМЫ   ================================}
{==============================================================================}
procedure TFKOD_CORRECT.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;


{==============================================================================}
{=====================   ВНЕШНИЙ ВЫЗОВ: ВЫПОЛНИТЬ    ==========================}
{==============================================================================}
procedure TFKOD_CORRECT.Execute(const STab, SCol, SRow: String);
begin
    If STab <> '' then CBFormula.Text := GeneratKod(STab, SRow, SCol)
                  else CBFormula.Text := '';
    EditChange(nil);
    ShowModal;
end;


{==============================================================================}
{=======================   СОБЫТИЕ: ИЗМЕНЕНИЕ ФОРМУЛЫ   =======================}
{==============================================================================}
procedure TFKOD_CORRECT.EditChange(Sender: TObject);
const CAPT = 'Коррекция кода';
      OPER = '+-*/()';
var I             : Integer;
    S1, S2, SCapt : String;
begin
    {Инициализация}
    CBFormula.OnChange := nil;
    ETab2.OnChange    := nil;
    ECol2.OnChange    := nil;
    ERow2.OnChange    := nil;
    IsFormula         := false;
    IsReplace         := false;
    IsDelTab          := false;
    IsDelKey          := false;
    try
       {Корректировка полей ввода}
       If Not IsIntegerStr(ETab2.Text) then ETab2.Text := '' else ETab2.Text := Trim(ETab2.Text);
       If Not IsIntegerStr(ECol2.Text) then ECol2.Text := '' else ECol2.Text := Trim(ECol2.Text);
       If Not IsIntegerStr(ERow2.Text) then ERow2.Text := '' else ERow2.Text := Trim(ERow2.Text);
       CBFormula.Text := Trim(CBFormula.Text);

       {Инициализация переменных}
       SFormula := CBFormula.Text;
       SeparatKod(SFormula, STab1, SRow1, SCol1);
       STab2 := ETab2.Text;
       SCol2 := ECol2.Text;
       SRow2 := ERow2.Text;

       {Это формула или нет}
       IsFormula := (STab2<>'') and (SRow2<>'') and (SCol2<>'');
       {Проверяем коды}
       If IsFormula then begin
          S1 := SFormula;
          S2 := CutModulChar(S1, '[', ']');
          While S2 <> '' do begin
             IsFormula := IsFormula and IsKod(S2);
             S1        := ReplModulChar(S1, '',  '[', ']');
             S2        := CutModulChar (S1, '[', ']');
          end;
       end;
       {Проверяем операции}
       If IsFormula then begin
          For I:=1 to Length(S1) do begin
             IsFormula := IsFormula and ((AnsiPos(S1[I], OPER) > 0) or IsIntegerStr(S1[I]));
          end;
          If SFormula <> '' then begin
             IsFormula := IsFormula and ((SFormula[Length(SFormula)]=']') or (SFormula[Length(SFormula)]=')') or IsIntegerStr(SFormula[Length(SFormula)]));
             IsFormula := IsFormula and ((SFormula[1]='[') or (SFormula[1]='(') or IsIntegerStr(SFormula[1]));
          end;
       end;

       IsReplace := (STab1<>'') and (SRow1<>'') and (SCol1<>'') and
                    (STab2<>'') and (SRow2<>'') and (SCol2<>'') and
                    IsIntegerStr(STab1) and IsIntegerStr(SRow1) and IsIntegerStr(SCol1) and
                    IsIntegerStr(STab2) and IsIntegerStr(SRow2) and IsIntegerStr(SCol2);
       IsDelTab  := (STab1<>'') and (SRow1='')  and (SCol1='')  and
                    (STab2='')  and (SRow2='')  and (SCol2='')  and
                    IsIntegerStr(STab1);
       IsDelKey  := (STab1<>'') and (SRow1<>'') and (SCol1<>'') and
                    (STab2='')  and (SRow2='')  and (SCol2='')  and
                    IsIntegerStr(STab1) and IsIntegerStr(SRow1) and IsIntegerStr(SCol1);

       {Заголовок и доступность Action}
       SCapt := CAPT;
       If IsFormula then SCapt := CAPT+': вычисление '       + SFormula + ' в ['+ GeneratKod(STab2, SRow2, SCol2)+']';
       If IsReplace then SCapt := CAPT+': замена '           + GeneratKod(STab1, SRow1, SCol1) + ' на ' + GeneratKod(STab2, SRow2, SCol2);
       If IsDelTab  then SCapt := CAPT+': удаление таблицы ' + STab1;
       If IsDelKey  then SCapt := CAPT+': удаление кода '    + GeneratKod(STab1, SRow1, SCol1);
       If (Not IsFormula) and (Not IsReplace) and (Not IsDelTab) and (Not IsDelKey) then SCapt := CAPT+' недопустима';
       Caption := SCapt;
    finally
       ARun.Enabled       := IsFormula or IsReplace or IsDelTab or IsDelKey;
       CBFormula.OnChange := EditChange;
       ETab2.OnChange     := EditChange;
       ECol2.OnChange     := EditChange;
       ERow2.OnChange     := EditChange;
    end;
end;


{==============================================================================}
{========================   ACTION: ВЫБРАТЬ ФОРМУЛУ   =========================}
{==============================================================================}
procedure TFKOD_CORRECT.AKodFormulaExecute(Sender: TObject);
var F : TFKOD_FORMULA;
begin
    F:=TFKOD_FORMULA.Create(Self);
    try     CBFormula.Text:=F.Execute(CBFormula.Text);
    finally F.Free;
    end;
end;


{==============================================================================}
{======================   ACTION: ВЫБРАТЬ СТАРЫЙ КОД   ========================}
{==============================================================================}
procedure TFKOD_CORRECT.AKodOldExecute(Sender: TObject);
var F : TFKOD;
    SKod, STab, SCol, SRow : String;
begin
    {Инициализация}
    SKod := CBFormula.Text;

    {Диалог}
    F:=TFKOD.Create(Self);
    try     SKod:=F.Execute(SKod, false);
    finally F.Free;
    end;

    {Меняем код}
    If Not SeparatKod(SKod, STab, SRow, SCol) then Exit;
    CBFormula.Text := SKod;
end;


{==============================================================================}
{======================   ACTION: ВЫБРАТЬ НОВЫЙ КОД   =========================}
{==============================================================================}
procedure TFKOD_CORRECT.AKodNewExecute(Sender: TObject);
var F : TFKOD;
    SKod, STab, SCol, SRow : String;
begin
    {Инициализация}
    SKod := GeneratKod(ETab2.Text, ERow2.Text, ECol2.Text);

    {Диалог}
    F:=TFKOD.Create(Self);
    try     SKod:=F.Execute(SKod, false);
    finally F.Free;
    end;

    {Меняем код}
    If Not SeparatKod(SKod, STab, SRow, SCol) then Exit;
    ETab2.Text := STab;
    ERow2.Text := SRow;
    ECol2.Text := SCol;
end;

procedure TFKOD_CORRECT.ARunExecute(Sender: TObject);
var S, SPath : String;
    I        : Integer;
begin
    {Корректируем и сохраняем список поиска}
    S := AnsiUpperCase(Trim(CBFormula.Text));
    If S = '' then Exit;
    I := CBFormula.Items.IndexOf(S);
    If I>=0 then CBFormula.Items.Delete(I);
    CBFormula.Items.Insert(0, S);
    CBFormula.Text := S;
    For I:=CBFormula.Items.Count-1 downto 25 do CBFormula.Items.Delete(I);
    WriteCBListIni(@CBFormula, INI_KOD_CORRECT);

    {Подтверждение}
    If MessageDlg('Будьте внимательны: отмена изменений будет невозможна!'+CH_NEW+
                  'Перед началом операции рекомендуется создать резервную копию базы данных.'+CH_NEW+CH_NEW+
                  'Подтвердите выполнение операции.',
                   mtWarning, [mbYes, mbNo], 0)<>mrYes then Exit;

    If IsDelTab then begin
       Case MessageDlg('Удалять только пустые таблицы?'+CH_NEW+CH_NEW+
                       'Да - удалять только пустые таблицы'+CH_NEW+
                       'Нет - удалять все таблицы',
                        mtConfirmation, [mbYes, mbNo], 0) of
          mrYes: IsDelTabEmpty := true;
          mrNo:  IsDelTabEmpty := false;
          else Exit;
       end;
    end;

    SPath := PATH_BD_DATA;
    If SPath[Length(SPath)]='\' then Delete(SPath, Length(SPath), 1);
    ARun.Enabled        := false;
    ACancel.Enabled     := false;
    AKodFormula.Enabled := false;
    AKodOld.Enabled     := false;
    AKodNew.Enabled     := false;
    POld.Enabled        := false;
    PNew.Enabled        := false;
    try     LoopDir(SPath);
            FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' Готово';
    finally Beep;
            ACancel.Enabled     := true;
            AKodFormula.Enabled := true;
            AKodOld.Enabled     := true;
            AKodNew.Enabled     := true;
            POld.Enabled        := true;
            PNew.Enabled        := true;
            EditChange(Sender);
    end;
end;


{==============================================================================}
{===========================   СОБЫТИЕ: ЗАКРЫТЬ    ============================}
{==============================================================================}
procedure TFKOD_CORRECT.ACancelExecute(Sender: TObject);
begin
    Close;
end;


{==============================================================================}
{===========   ВЛОЖЕННЫЙ ПРОСМОТР ВСЕХ ФАЙЛОВ В КАТОЛАГЕ SPATH    =============}
{==============================================================================}
procedure TFKOD_CORRECT.LoopDir(const SPath: String);
var Sr : TSearchRec;
    S  : String;
begin
    try
       If FindFirst(SPath+'\*.*', faAnyFile, Sr)= 0 then begin
          repeat
             S:=Sr.Name;
             If (S='.') or (S='..') then Continue;
             If (Sr.Attr and faDirectory) > 0 then LoopDir(SPath+'\'+S)
                                              else If CmpStr(ExtractFileExt(S), '.dat')
                                                   then StepFile(SPath+'\'+S);
             Application.ProcessMessages;
          until FindNext(Sr) <> 0;
       end;
    finally
       FindClose(Sr);
    end;
end;



{==============================================================================}
{==============================   ЗАМЕНА КОДА   ===============================}
{==============================================================================}
function TFKOD_CORRECT.StepFile(const SFile: String): Boolean;
const SLEN = 30;
var FDat                       : TMemIniFile;
    SVal, SKey1, SKey2, S1, S2 : String;
    SList                      : TStringList;
begin
    {Инициализация}
    Result := false;
    {Информатор}
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' '+SFile;
    SVal  := SFile;
    SKey1 := SCol1+' '+SRow1;
    SKey2 := SCol2+' '+SRow2;

    FDat := TMemIniFile.Create(SFile);
    try
       {Удаление таблиц}
       If IsDelTab  then begin
          {Удаление только пустых таблиц}
          If IsDelTabEmpty then begin
             SList:=TStringList.Create;
             try     FDat.ReadSection(STab1, SList);
                     If SList.Count = 0 then FDat.EraseSection(STab1);
             finally SList.Free;
             end;
          {Удаление всех таблиц}
          end else begin
             FDat.EraseSection(STab1);
          end;
       end;

       {Удаление ключей}
       If IsDelKey  then FDat.DeleteKey(STab1, SKey1);

       {Выражение}
       If IsFormula then begin
          {Меняем коды на значения}
          S1 := SFormula;
          S2 := CutModulChar(S1, '[', ']');
          While S2 <> '' do begin
             SeparatKod(S2, STab1, SRow1, SCol1);
             SVal := FDat.ReadString(STab1, SCol1+' '+SRow1, '');
             If SVal='' then SVal:='0';
             S1 := ReplModulChar(S1, SVal,  '[', ']');
             S2 := CutModulChar (S1, '[', ']');
          end;
          {Вычисляем выражение}
          SVal := FloatToStr(CalcStr(S1));
          {Записываем значение, пустые ключи удаляем}
          If (SVal='0') or (SVal='') then FDat.DeleteKey(STab2, SKey2)
                                     else FDat.WriteString(STab2, SKey2, SVal);
       end;

       {Замена}
       If IsReplace then begin
          {Читаем и удаляем старое значение}
          SVal := FDat.ReadString(STab1, SKey1, '');
          FDat.DeleteKey(STab1, SKey1);
          {Записываем значение, пустые ключи удаляем}
          If (SVal='0') or (SVal='') then FDat.DeleteKey(STab2, SKey2)
                                     else FDat.WriteString(STab2, SKey2, SVal);
       end;

       {Сохраняем изменения}
       FDat.UpdateFile;
       {Возвращаемый результат}
       Result:=true;
    finally
       FDat.Free;
    end;
end;


end.
