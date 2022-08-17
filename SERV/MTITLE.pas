unit MTITLE;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.ActnList, Vcl.ExtCtrls, Vcl.Buttons,
  Data.DB, Data.Win.ADODB,
  MAIN;

type
  TFTITLE = class(TForm)
    ActionList1: TActionList;
    AOK: TAction;
    ACancel: TAction;
    ARestore: TAction;
    Bevel1: TBevel;
    ASave: TAction;
    Panel1: TPanel;
    SpeedButton1: TSpeedButton;
    BtnCancel: TBitBtn;
    BtnOk: TBitBtn;
    Panel2: TPanel;
    Panel3: TPanel;
    BtnRestore: TBitBtn;
    PPeriod: TPanel;
    LPeriod: TLabel;
    CBYear: TComboBox;
    LPeriodSpace: TPanel;
    CBMonth: TComboBox;
    PForm: TPanel;
    LForm: TLabel;
    CBForm: TComboBox;
    PRegion: TPanel;
    LReg: TLabel;
    CBRegion: TComboBox;
    PSubRegion: TPanel;
    LSubRegion: TLabel;
    CBSubRegion: TEdit;
    CBPrev: TCheckBox;
    procedure AOKExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CBPeriodChange(Sender: TObject);
    procedure CBRegionChange(Sender: TObject);
    procedure CBFormChange(Sender: TObject);
    procedure ARestoreExecute(Sender: TObject);
    procedure CBSubRegionChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormHide(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure ASaveExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FFMAIN : TFMAIN;
    SYear0, SMonth0, SRegion0, SRegion01, SRegion02, SForm0: String;
    LIni   : TStringList;

    procedure EnablAction;
  public
    function  Execute(var SYear, SMonth, SRegion, SForm: String;
                      const IsSelForm: Boolean;
                      const IsDialog: Boolean): Boolean;
  end;

var
  FTITLE : TFTITLE;

implementation

uses FunConst, FunSys, FunDay, FunText, FunIO;

{$R *.dfm}


{==============================================================================}
{==========================   �������� �����   ================================}
{==============================================================================}
procedure TFTITLE.FormCreate(Sender: TObject);
var I: Integer;
begin
    {�������������}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' �������� �� ������';
    LIni   := TStringList.Create;

    CBYear.OnChange            := CBPeriodChange;
    CBYear.OnKeyDown           := FormKeyDown;

    CBMonth.AutoComplete       := true;
    CBMonth.OnChange           := CBPeriodChange;
    CBMonth.OnKeyDown          := FormKeyDown;

    CBForm.AutoComplete        := true;
    CBForm.Style               := csDropDown;
    CBForm.OnChange            := CBFormChange;
    CBForm.OnKeyDown           := FormKeyDown;

    CBRegion.AutoComplete      := true;
    CBRegion.Style             := csDropDown;
    CBRegion.OnChange          := CBRegionChange;
    CBRegion.OnKeyDown         := FormKeyDown;

    CBSubRegion.OnChange       := CBSubRegionChange;
    CBSubRegion.OnKeyDown      := FormKeyDown;

    {������������� ������ ���}
    For I:=CurrentYear-3 to CurrentYear do CBYear.Items.Add(IntToStr(I));

    {������������� ������ �������}
    For I:=Low(MonthList) to High(MonthList) do CBMonth.Items.Add(MonthList[I]);
end;


{==============================================================================}
{============================   ����� �����   =================================}
{==============================================================================}
procedure TFTITLE.FormShow(Sender: TObject);
begin
    CBYear.SelLength      := 0;
    CBMonth.SelLength     := 0;
    CBForm.SelLength      := 0;
    CBRegion.SelLength    := 0;
    CBSubRegion.SelLength := 0;
end;


{==============================================================================}
{===========================   ������� �����   ================================}
{==============================================================================}
procedure TFTITLE.FormHide(Sender: TObject);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;


{==============================================================================}
{===========================   �������� �����   ===============================}
{==============================================================================}
procedure TFTITLE.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    LIni.Free;
end;


{==============================================================================}
{=====================   ������� �����: ���������    ==========================}
{==============================================================================}
function TFTITLE.Execute(var SYear, SMonth, SRegion, SForm: String;
                         const IsSelForm: Boolean;          // �������� �� ����� �����
                         const IsDialog: Boolean): Boolean; // ���������� �� ������ (��� ��� �������)
const STR_PREV = '�������� ��������: ';
var Q : TADOQuery;
    S : String;
begin
    {�������������}
    Result     := false;
    SYear0     := SYear;
    SMonth0    := SMonth;
    SRegion0   := SRegion;
    SRegion01  := ReplModulStr(SRegion0, '', SUB_RGN1, SUB_RGN2);
    SRegion02  := CutModulStr(SRegion0,      SUB_RGN1, SUB_RGN2);
    SForm0     := SForm;
    
    {*** ����������� ������� ������ �� ���������� ������ **********************}
    If IsSelForm then begin
       //CBPrev.Checked := true;
    end else begin
       CBPrev.Visible := false;
       Height := Height - CBPrev.Height;  
    end;

    {*** ������������� ������ *************************************************}
    CBYear.Text      := SYear;
    CBMonth.Text     := SMonth;

    {*** ������������� ������ *************************************************}
    CBSubRegion.Text := SRegion02;
    CBRegion.Text    := SRegion01;

    {*** ������� ������� ******************************************************}
    If IsDialog then begin
       Q:=TADOQuery.Create(Self);
       try
          Q.Connection:=FFMAIN.LBank[0].BD;
          Q.SQL.Text  := 'SELECT * FROM ['+TABLE_REGREPL+'] '+
                         'WHERE (['+REGREPL_R1+'] = '+QuotedStr(CBRegion.Text)+') '+
                         'AND   ((['+REGREPL_S1+'] = '+QuotedStr(CBSubRegion.Text)+')'+
                         '   OR (('+QuotedStr(CBSubRegion.Text)+' = '+QuotedStr('')+') AND (['+REGREPL_S1+'] Is Null)));';
          Q.Open;
          If Q.RecordCount = 1 then begin
             CBRegion.Text    := Q.FieldByName(REGREPL_R2).AsString;
             CBSubRegion.Text := Q.FieldByName(REGREPL_S2).AsString;
          end;
       finally
          If Q.Active then Q.Close;
          Q.Free;
       end;
    end;

    {*** ������������� ����� **************************************************}
    CBForm.Enabled := IsSelForm;
    If IsSelForm then CBForm.Color := clWindow else CBForm.Color := clBtnFace;
    CBForm.Text    := SForm;

    {��������� ������}
    CBPeriodChange(nil);

    {��������� Hint}
    CBYear.Hint      := STR_PREV + SYear0;
    CBMonth.Hint     := STR_PREV + SMonth0;
    CBForm.Hint      := STR_PREV + SForm0;
    CBSubRegion.Hint := STR_PREV + SRegion02;
    CBRegion.Hint    := STR_PREV + SRegion01;

    {���������� ����}
    If IsDialog or (Not AOK.Enabled) then begin
       If (ShowModal<>mrYes) then Exit;
    end;

    {������������ ���������}
    Result  := true;
    SYear   := CBYear.Text;
    SMonth  := CBMonth.Text;
    If IsSelForm then SForm:=CBForm.Text;
    SRegion := CBRegion.Text;
    S       := Trim(CBSubRegion.Text);
    If S<>'' then SRegion := SRegion+SUB_RGN1+S+SUB_RGN2;
end;


{==============================================================================}
{========================   ��������� �������   ===============================}
{==============================================================================}
procedure TFTITLE.CBPeriodChange(Sender: TObject);
label Er, Dl;
var DatPeriod      : TDate;
    LRegion_       : TLRegion;
    LForma_        : TLForma;
    IYear, IMonth  : Word;
    SRegion, SForm : String;
    SFile          : String;
    I              : Integer;
begin
    {�������������}
    IYear   := 0;
    SRegion := CBRegion.Text;
    SForm   := CBForm.Text;
    CBYear.Items.BeginUpdate;
    CBMonth.Items.BeginUpdate;
    CBRegion.Items.BeginUpdate;
    CBForm.Items.BeginUpdate;
    try
       CBRegion.Items.Clear;
       CBForm.Items.Clear;
       LIni.Clear;

       {������������� ���}
       If IsIntegerStr(CBYear.Text) then begin
          IYear :=StrToInt(CBYear.Text);
          If (IYear>2100) or (IYear<1900) then CBYear.Font.Color := clRed
                                          else CBYear.Font.Color := clBlack;
       end else begin
          CBYear.Font.Color := clRed;
       end;

       {������������� �����}
       IMonth := MonthStrToInd(CBMonth.Text);
       If IMonth>0 then CBMonth.Font.Color := clBlack
                   else CBMonth.Font.Color := clRed;

       {������������ �������}
       If (CBYear.Font.Color=clRed) or (CBMonth.Font.Color=clRed) then Goto Er;
       DatPeriod := IDatePeriod(IYear, IMonth);
       If DatPeriod=0 then Goto Er;

       {������������� ������ ��������}
       If Not FFMAIN.FindRegions(@LRegion_, false, -1, '', -1, -1, DatPeriod) then Goto Er;
       For I:=Low(LRegion_) to High(LRegion_) do CBRegion.Items.Add(LRegion_[I].FCaption);

       {������������� �����}
       If Not FFMAIN.FindForms(@LForma_, false, true, -1, -1, '', '', '', CBMonth.Text, false, DatPeriod) then Goto Er;
       For I:=Low(LForma_) to High(LForma_) do begin
          If (LForma_[I].FCaption='') or (LForma_[I].FIcon=0) then Continue;
          {�������� ��������� ��� ����������� ������}
          SFile := FFMAIN.PathFileForm(LForma_[I].SBankPref, LForma_[I].FFile);
          If Not FFMAIN.IsFormReadOnly(LForma_[I].FCounter) then begin
             CBForm.Items.Add(LForma_[I].FCaption);
             LIni.Add(ChangeFileExt(SFile, '.ini'));
          end;
       end;

       Goto Dl;

Er:    {��������� ������}
       // SRegion:='';
       If CBForm.Enabled then SForm:='';

    {���������� ���������}   
Dl: finally
       SetLength(LRegion_, 0);
       SetLength(LForma_,  0);

       CBRegion.Text := SRegion;
//       CBRegionChange(Sender);
//       CBRegion.Items.EndUpdate;

       CBForm.Text := SForm;
       CBFormChange(Sender);
       CBForm.Items.EndUpdate;
       CBRegion.Items.EndUpdate;

       CBMonth.Items.EndUpdate;
       CBYear.Items.EndUpdate;
    end;
end;


{==============================================================================}
{==========================   ��������� �����   ===============================}
{==============================================================================}
procedure TFTITLE.CBFormChange(Sender: TObject);
begin
    {�������� �����}
    If CBForm.Items.IndexOf(CBForm.Text)>=0 then CBForm.Font.Color := clBlack
                                            else CBForm.Font.Color := clRed;

    {���������}
    CBRegionChange(Sender);
    // {����������� Action}
    // EnablAction;
end;


{==============================================================================}
{========================   ��������� �������   ===============================}
{==============================================================================}
procedure TFTITLE.CBRegionChange(Sender: TObject);
var LTab : TStringList;
    Ind  : Integer;
    IsOk : Boolean;
begin
    {�������������}
    Ind := CBForm.Items.IndexOf(CBForm.Text);
    If (Ind >= 0) and (Ind <= (LIni.Count-1)) then begin
       LTab:=TStringList.Create;
       {�������� ������������}
       try
          {������ ������ ������}
          If Not FFMAIN.GetListTabIni(@LTab, LIni[Ind]) then LTab.Clear;
          {�������� �����������}
          IsOk:=FFMAIN.IsModifyRegion(CBYear.Text, CBMonth.Text, CBRegion.Text+SUB_RGN1+CBSubRegion.Text+SUB_RGN2, @LTab);
          {�������� ������������ �������� �������}
          If IsOk then IsOk:=(CBRegion.Items.IndexOf(CBRegion.Text)>=0);
       finally
          LTab.Free;
       end;
    end else begin
       IsOk:=false;
    end;

    {���������}
    CBRegion.Items.BeginUpdate;
    try
       If IsOk then begin
          CBRegion.Font.Color    := clBlack;
          CBSubRegion.Font.Color := clBlack;
       end else begin
          CBRegion.Font.Color    := clRed;
          CBSubRegion.Font.Color := clRed;
       end;
    finally
       CBRegion.Items.EndUpdate;
    end;

    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{=====================   ������������ ������� ����������   ====================}
{==============================================================================}
procedure TFTITLE.CBSubRegionChange(Sender: TObject);
begin
    CBRegionChange(Sender);
end;


{==============================================================================}
{============================   ������� �������   =============================}
{==============================================================================}
procedure TFTITLE.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    Case Key of
    13: {ENTER} If AOK.Enabled then AOK.Execute;
    27: {ESC}   ModalResult:=mrCancel;
    end;
end;


{==============================================================================}
{=====================   ���������� ���������� �������   ======================}
{==============================================================================}
procedure TFTITLE.ASaveExecute(Sender: TObject);
var T : TADOTable;
    S : String;
begin
    {�������������}
    If IS_NET then Exit;
    If Trim(CBRegion.Text) = '' then Exit;
    If Trim(SRegion0)      = '' then Exit;
    If (CBRegion.Text=SRegion01) and (CBSubRegion.Text=SRegion02) then Exit;

    T:=TADOTable.Create(Self);
    try
       T.Connection := FFMAIN.LBank[0].BD;
       T.TableName  := TABLE_REGREPL;
       If SRegion02 <> ''
       then S :=  '(['+REGREPL_R1+']='+QuotedStr(SRegion01)+') AND (['+REGREPL_S1+']='+QuotedStr(SRegion02)+')'
       else S := '((['+REGREPL_R1+']='+QuotedStr(SRegion01)+') AND (['+REGREPL_S1+']='+QuotedStr(SRegion02)+')) OR '+
                 '((['+REGREPL_R1+']='+QuotedStr(SRegion01)+') AND (['+REGREPL_S1+']=Null))'; // �����������: (A or B) and C    ���������: (A and C) or (B and C)
       T.Filter     := S;
       T.Filtered   := true;
       T.Open;

       {������� ��� ������}
       If T.RecordCount > 0 then begin
          If (CBRegion.Text=T.FieldByName(REGREPL_R2).AsString) and (CBSubRegion.Text=T.FieldByName(REGREPL_S2).AsString) then Exit;
          T.Edit;
       end else begin
          T.Insert;
          T.FieldByName(REGREPL_R1).AsString := SRegion01;
          T.FieldByName(REGREPL_S1).AsString := SRegion02;
       end;

       {������ ���������}
       T.FieldByName(REGREPL_R2).AsString := CBRegion.Text;
       T.FieldByName(REGREPL_S2).AsString := CBSubRegion.Text;
       T.Post;

       {�����������}
       MessageDlg('������� ���������� ����������!'+CH_NEW+
                   SRegion0+CH_NEW+
                   '������ ��'+CH_NEW+
                   CBRegion.Text+' '+CBSubRegion.Text,
                   mtInformation, [mbOk], 0);
    finally
       If T.Active then T.Close;
       T.Free;
    end;
end;

{==============================================================================}
{========================   ����������� ACTION   ==============================}
{==============================================================================}
procedure TFTITLE.EnablAction;
var B_OK, B_Restore : Boolean;
    BYear, BMonth, BForm, BRegion, BSubRegion, BPrev : Boolean;
begin
    {�������������}
    B_OK:=true;

    {��������}
    If (CBYear.Font.Color      = clRed) or
       (CBMonth.Font.Color     = clRed) or
       (CBForm.Font.Color      = clRed) or
       (CBRegion.Font.Color    = clRed) or
       (CBSubRegion.Font.Color = clRed) or
       (FindStr(SUB_RGN1, CBSubRegion.Text)>0) or
       (FindStr(SUB_RGN2, CBSubRegion.Text)>0) then B_OK:=false;

    {��������������}
    BYear      := CBYear.Text      <> SYear0;
    BMonth     := CBMonth.Text     <> SMonth0;
    BForm      := CBForm.Text      <> SForm0;
    BSubRegion := CBSubRegion.Text <> SRegion02;
    BRegion    := CBRegion.Text    <> SRegion01;

    B_Restore:= BYear or BMonth or BForm or BSubRegion or BRegion;

    CBYear.ShowHint      := BYear;
    CBMonth.ShowHint     := BMonth;
    CBForm.ShowHint      := BForm;
    CBSubRegion.ShowHint := BSubRegion;
    CBRegion.ShowHint    := BRegion;

    If BYear      then CBYear.Color      := $00E1FFFF else CBYear.Color      := clWhite;
    If BMonth     then CBMonth.Color     := $00E1FFFF else CBMonth.Color     := clWhite;
    If CBForm.Color<>clBtnFace then begin If BForm    then CBForm.Color      := $00E1FFFF else CBForm.Color := clWhite; end;
    If BSubRegion then CBSubRegion.Color := $00E1FFFF else CBSubRegion.Color := clWhite;
    If BRegion    then CBRegion.Color    := $00E1FFFF else CBRegion.Color    := clWhite;

    {�����������}
    AOK.Enabled      := B_OK;
    ARestore.Enabled := B_Restore;

    {���������� ������� ������ ���: ���������� ������������ � ����������� ��������������� �������}
    BPrev := true;
    If (MonthStrToInd(CBMonth.Text) = 1) and (FFMAIN.IsFormPeriod(CBForm.Text, CBYear.Text, CBMonth.Text, true,  false)) then BPrev := false;
    If (MonthStrToInd(CBMonth.Text) = 3) and (FFMAIN.IsFormPeriod(CBForm.Text, CBYear.Text, CBMonth.Text, false, true))  then BPrev := false;
    CBPrev.Enabled := BPrev;
    If Not BPrev then CBPrev.Checked := false;
end;


procedure TFTITLE.AOKExecute(Sender: TObject);
begin
    ASaveExecute(Sender);
    ModalResult:=mrYes;
end;


procedure TFTITLE.ACancelExecute(Sender: TObject);
begin
    ModalResult:=mrNo;
end;


procedure TFTITLE.ARestoreExecute(Sender: TObject);
begin
    {������������� ������ � ������}
    CBYear.Text      := SYear0;
    CBMonth.Text     := SMonth0;
    CBForm.Text      := SForm0;
    CBSubRegion.Text := SRegion02;
    CBRegion.Text    := SRegion01;
    CBPeriodChange(nil);
end;

end.
