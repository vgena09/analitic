unit MKOD_FORMULA;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls,
  Vcl.Menus, Vcl.ActnList, Vcl.StdCtrls, Vcl.Buttons,
  MAIN, Vcl.ComCtrls, Vcl.ToolWin;

type
  TFKOD_FORMULA = class(TForm)
    PBottom: TPanel;
    PFormula: TPanel;
    Bevel1: TBevel;
    ActionList1: TActionList;
    BtnOk: TBitBtn;
    BtnCancel: TBitBtn;
    Bevel3: TBevel;
    LFormula: TLabel;
    CBFormula: TComboBox;
    Bevel2: TBevel;
    AOk: TAction;
    CoolBar1: TCoolBar;
    ToolBar1: TToolBar;
    BtnOpenStandart: TToolButton;
    BtnOpen: TToolButton;
    Separator1: TToolButton;
    BtnSave: TToolButton;
    BtnClear: TToolButton;
    PDesc: TPanel;
    EDesc: TRichEdit;
    LDesc: TLabel;
    AOpen: TAction;
    AOpenStandart: TAction;
    ASave: TAction;
    AClear: TAction;
    OpenDlg: TOpenDialog;
    SaveDlg: TSaveDialog;
    Bevel4: TBevel;
    AKod: TAction;
    BtnAddKod: TToolButton;
    ToolButton1: TToolButton;
    procedure FormCreate(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure CBFormulaChange(Sender: TObject);
    procedure AOkExecute(Sender: TObject);
    procedure ASaveExecute(Sender: TObject);
    procedure AOpenExecute(Sender: TObject);
    procedure AOpenStandartExecute(Sender: TObject);
    procedure AClearExecute(Sender: TObject);
    procedure CBFormulaKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure PDescResize(Sender: TObject);
    procedure AKodExecute(Sender: TObject);
  private
    FFMAIN          : TFMAIN;
    CONFIG_AUTOSAVE : String;

    procedure EnablAction(const IsOk: Boolean);
    function  SetDesc: Boolean;
  public
    function  Execute(const SFormula : String): String;
  end;

const MAX_CBLIST = 50;
      CONFIG_EXT = 'formula';

var
  FKOD_FORMULA: TFKOD_FORMULA;

implementation

uses MKOD, FunConst, FunText, FunSys, FunFiles, FunIni;

{$R *.dfm}


{==============================================================================}
{=============================   �������� �����    ============================}
{==============================================================================}
procedure TFKOD_FORMULA.FormCreate(Sender: TObject);
begin
    {�������������}
    FFMAIN:= TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' �������� �������';
    BtnCancel.ModalResult := mrCancel;
    EDesc.ReadOnly        := true;

    CONFIG_AUTOSAVE    := '��������������.'+CONFIG_EXT;
    SaveDlg.DefaultExt := CONFIG_EXT;
    SaveDlg.Filter     := '����� ������|*.'+CONFIG_EXT+'|��� �����|*.*';
    SaveDlg.Title      := '���������� ������';
    OpenDlg.DefaultExt := CONFIG_EXT;
    OpenDlg.Filter     := '����� ������|*.'+CONFIG_EXT+'|��� �����|*.*';
    OpenDlg.Title      := '�������������� ������';

    {��������������� �������}
    CBFormula.DropDownCount := MAX_CBLIST;
    ReadCBListIni(@CBFormula, INI_KOD_FORMULA);
end;


{==============================================================================}
{==============================   ����� �����    ============================}
{==============================================================================}
procedure TFKOD_FORMULA.FormShow(Sender: TObject);
begin
    CBFormula.SetFocus;
end;

procedure TFKOD_FORMULA.PDescResize(Sender: TObject);
begin
    EDesc.Repaint;
end;

{==============================================================================}
{==============================   ������� �����    ============================}
{==============================================================================}
procedure TFKOD_FORMULA.FormHide(Sender: TObject);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;


{==============================================================================}
{=====================   ������� �����: ���������    ==========================}
{==============================================================================}
function TFKOD_FORMULA.Execute(const SFormula : String): String;
begin
    {����������� �������}
    CBFormula.Text := SFormula;
    SetDesc;

    {���������� �����}
    If ShowModal = mrOk then Result:=Trim(CBFormula.Text)
                        else Result:=SFormula;
end;


{==============================================================================}
{===========================   ����������� ACTION    ==========================}
{==============================================================================}
procedure TFKOD_FORMULA.EnablAction(const IsOk: Boolean);
begin
    AOk.Enabled := IsOk;
end;


{==============================================================================}
{=========================   ACTION: ���������   ==============================}
{==============================================================================}
procedure TFKOD_FORMULA.ASaveExecute(Sender: TObject);
var FName: String;
begin
    SaveDlg.InitialDir := PATH_WORK;
    If Not SaveDlg.Execute then Exit;
    FName := SaveDlg.FileName;
    WriteTxtFile(FName, CBFormula.Text);
end;


{==============================================================================}
{=======================   ACTION: ������������   =============================}
{==============================================================================}
procedure TFKOD_FORMULA.AOpenExecute(Sender: TObject);
var FName : String;
begin
    OpenDlg.InitialDir := PATH_WORK;
    OpenDlg.FileName   := '';
    If Not OpenDlg.Execute then Exit;
    FName          := OpenDlg.FileName;
    CBFormula.Text := ReadTxtFile(FName);
    SetDesc;
end;

procedure TFKOD_FORMULA.AOpenStandartExecute(Sender: TObject);
var FName : String;
begin
    OpenDlg.InitialDir := PATH_BD_SHABLON;
    OpenDlg.FileName   := '';
    If Not OpenDlg.Execute then Exit;
    FName          := OpenDlg.FileName;
    CBFormula.Text := ReadTxtFile(FName);
    SetDesc;
end;


{==============================================================================}
{==========================   ACTION: ��������   ==============================}
{==============================================================================}
procedure TFKOD_FORMULA.AClearExecute(Sender: TObject);
begin
    CBFormula.Text:='';
    SetDesc;
end;


{==============================================================================}
{========================   ACTION: �������� ���   ============================}
{==============================================================================}
procedure TFKOD_FORMULA.AKodExecute(Sender: TObject);
var F            : TFKOD;
    S, SSel      : String;
    IStart, ILen : Integer;
begin
    IStart := CBFormula.SelStart;
    ILen   := CBFormula.SelLength;
    S      := CutModulChar(CBFormula.SelText, '[', ']');
    If S='' then SSel := CBFormula.SelText
            else SSel := S;
    F      := TFKOD.Create(Self);
    try S:=F.Execute(SSel, false);
        CBFormula.SetFocus;
        CBFormula.SelStart  := IStart;
        CBFormula.SelLength := ILen;
        If S<>SSel then CBFormula.SelText := '['+S+']';
    finally F.Free;
    end;
end;


{==============================================================================}
{=========================   ACTION: ����� �������   ==========================}
{==============================================================================}
procedure TFKOD_FORMULA.AOkExecute(Sender: TObject);
var S : String;
    I : Integer;
begin
    {������������ � ��������� ������ ������}
    S := AnsiUpperCase(Trim(CBFormula.Text));
    I := CBFormula.Items.IndexOf(S);
    If I>=0 then CBFormula.Items.Delete(I);
    CBFormula.Items.Insert(0, S);
    CBFormula.Text := S;
    For I:=CBFormula.Items.Count-1 downto MAX_CBLIST do CBFormula.Items.Delete(I);
    WriteCBListIni(@CBFormula, INI_KOD_FORMULA);

    {������������ ���������}
    ModalResult:=mrOk;
end;


{==============================================================================}
{=======================   �������: ��������� �������   =======================}
{==============================================================================}
procedure TFKOD_FORMULA.CBFormulaChange(Sender: TObject);
begin
    SetDesc;
end;


{==============================================================================}
{===================   �������: ENTER ��� ����� �������   =====================}
{==============================================================================}
procedure TFKOD_FORMULA.CBFormulaKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    Case Key of
    13: {ENTER} If AOK.Enabled then AOK.Execute;
    27: {ESC}   ModalResult:=mrCancel;
    end;
end;


{==============================================================================}
{==========================    �������� �������    ============================}
{==============================================================================}
function TFKOD_FORMULA.SetDesc: Boolean;
var SForm, SDesc, S : String;
    LKey            : TStringList;
    B               : Boolean;

   function AddColorText(const Text: String; const Color: TColor; const Style: TFontStyles): String;
   begin
       EDesc.SelAttributes.Color := Color;
       EDesc.SelAttributes.Style := Style;
       EDesc.SelText := Text;
       EDesc.SelAttributes.Color := clBlack;
   end;

   function AddDesc: Boolean;
   begin
       If SDesc='' then begin
          SDesc := '������������';
          Result:=false;
       end else begin
          Result:=true;
       end;
       If LKey.IndexOf(SForm) >= 0 then Exit;
       LKey.Add(SForm);
       AddColorText('['+SForm+']'+CH_NEW, clRed, [fsBold]);
       AddColorText(SDesc+CH_NEW, clBlue, []);
       AddColorText(CH_NEW, clBlack, []);
   end;

begin
    {�������������}
    Result := false;
    EDesc.Lines.BeginUpdate;
    LKey := TStringList.Create;
    try
       EDesc.Clear;
       LKey.Clear;
       S     := Trim(CBFormula.Text);
       SDesc := FFMAIN.KeyToCaption(S, CH_SPR, CH_NEW, false);
       If SDesc <> '' then begin
          {������ 1 ���}
          SForm  := S;
          Result := AddDesc;
       end else begin
          {��������� �����}
          SForm := CutModulChar(S, '[', ']');
          If SForm <> '' then Result:=true;
          While SForm <> '' do begin
             SDesc  := FFMAIN.KeyToCaption(SForm, CH_SPR, CH_NEW, false);
             B      := AddDesc;
             Result := Result and B;
             S      := ReplModulChar(S, '', '[', ']');
             SForm  := CutModulChar (S,     '[', ']');
          end;
       end;
    finally
       LKey.Free;
       EDesc.Lines.EndUpdate;
       EnablAction(Result);
    end;
end;

end.
