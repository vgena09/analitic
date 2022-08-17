unit MEXPORT;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Dialogs, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.FileCtrl,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ActnList, Vcl.Buttons,
  MAIN;

type
  TFEXPORT = class(TForm)
    ActionList: TActionList;
    AExport: TAction;
    ACancel: TAction;
    PPeriod: TPanel;
    LPeriod: TLabel;
    CBYear: TComboBox;
    CBMonth: TComboBox;
    PFolder: TPanel;
    LFolder: TLabel;
    PComment: TPanel;
    LComment: TLabel;
    CBComment: TMemo;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    PBottom: TPanel;
    Label1: TLabel;
    Panel4: TPanel;
    BtnCancel: TBitBtn;
    BtnOk: TBitBtn;
    CBFolder: TButtonedEdit;
    procedure FormCreate(Sender: TObject);
    procedure AExportExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure CBYearChange(Sender: TObject);
    procedure CBMonthChange(Sender: TObject);
    procedure CBFolderChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CBFolderRightButtonClick(Sender: TObject);
  private
    FFMAIN : TFMAIN;
    PasRar : String;

    procedure EnablAction;
    function  IsExistsAllRegions: Boolean;
  public
    function  Execute: Boolean;
  end;

var
  FEXPORT: TFEXPORT;

implementation

uses FunConst, FunText, FunDay, FunSys, FunInfo, FunVcl, FunFiles, FunBD, FunIni;

{$R *.dfm}


{==============================================================================}
{============================   �������� �����    =============================}
{==============================================================================}
procedure TFEXPORT.FormCreate(Sender: TObject);
var I : Integer;
    B : Boolean;
begin
    {�������������}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));

    {��������������� ��������� �� Ini}
    LoadFormIni(Self, [fspPosition]);

    {������������� ������ ���}
    For I:=CurrentYear-3 to CurrentYear do CBYear.Items.Add(IntToStr(I));

    {������������� ������ �������}
    CBMonth.Style := csDropDownList;
    For I:=Low(MonthList) to High(MonthList) do CBMonth.Items.Add(MonthList[I]);

    {������ ���� �������� ������ �� ����� ������� ��������}
    CBFolder.Text := ReadLocalString(INI_SET, INI_SET_EXPORT_PATH, '');

    {������ ������ �� ����� ���������� ��������}
    PasRar := ReadGlobalString(INI_SET, INI_SET_PAS_IMPORT_EXPORT, '');

    {����������� �����������}
    B := Not (Pos('-', PATH_WORK) > 0);
    LComment.Enabled  := B;
    CBComment.Enabled := B;
end;


{==============================================================================}
{============================   �������� �����    =============================}
{==============================================================================}
procedure TFEXPORT.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {��������� ��������� � Ini}
    SaveFormIni(Self, [fspPosition]);
end;


{==============================================================================}
{============================   ������� �������    ============================}
{==============================================================================}
function TFEXPORT.Execute: Boolean;
var SYear, SMonth: String;
begin
    {�������������}
    Result:=false;

    {���������� ������ �� ���������}
    If Not NullPeriod(SYear, SMonth) then Exit;
    CBYear.Text  := SYear;
    CBMonth.ItemIndex := CBMonth.Items.IndexOf(SMonth);

    {����������� Action}
    EnablAction;

    {����� �����}
    Result := (ShowModal = mrOk);
end;


{==============================================================================}
{==========================   ����������� ACTION   ============================}
{==============================================================================}
procedure TFEXPORT.EnablAction;
var B1, B2 : Boolean;
    IYear  : Integer;
begin
    {�������������}
    B1 := DirectoryExists(PATH_BD_DATA+CBYear.Text+'\'+CBMonth.Text);
    B2 := DirectoryExists(CBFolder.Text);
    AExport.Enabled := B1 and B2;

    {��������� ������}
    If IsIntegerStr(CBYear.Text) then begin
       IYear :=StrToInt(CBYear.Text);
       If (IYear>2100) or (IYear<1900) then B1 := false;
    end else B1:=false;

    If B1 then CBYear.Font.Color := clBlack
          else CBYear.Font.Color := clRed;
    CBMonth.Font.Color := CBYear.Font.Color;

    {��������� �������}
    If B2 then CBFolder.Font.Color := clBlack
          else CBFolder.Font.Color := clRed;
end;


{==============================================================================}
{============================   ACTION: �������   =============================}
{==============================================================================}
procedure TFEXPORT.AExportExecute(Sender: TObject);
var SRar     : String;
    SDect    : String;
    SComment : String;
    SKey     : String;
begin
    {�������������� ���� �� ��� �������}
    If Not IsExistsAllRegions then Exit;

    {�������������}
    Hide;
    FFMAIN.Repaint; 
    SComment := '';
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������� ������: ������';
    try
       {���� ����������}
       SRar   := GetProgramAssociation('.rar');
       If Not FileExists(SRar) then begin ErrMsg('�� ���������� ��������� ��� ������ ������ � RAR-�����!'); Exit; end;

       {���� ����������}
       SDect:=CBFolder.Text;
       If SDect[Length(SDect)] <> '\' then SDect:=SDect+'\';
       SDect:=SDect+'����������_'+CBYear.Text+'_'+CBMonth.Text;
       If SMAIN_REGION <> '' then SDect:=SDect+'_'+SMAIN_REGION;
       SDect:=SDect+'.rar';

       {������� ���������� ���� ����������}
       If FileExists(SDect) then begin
          If MessageDlg('���� � ����������� ������ ��� ����������.'+CH_NEW+'������� ������ ����?',
                        mtWarning, [mbYes, mbNo], 0)<>mrYes then Exit;
          DeleteFile(SDect);
       end;

       {��������� ����������� ���� �� ����}
       CBComment.Lines.Insert(0, ID_VERSION_PROG+GetProgFileVersion);
       If Trim(CBComment.Text) <> '' then begin
          SComment := PATH_WORK_TEMP+TEMP_COMMENT;
          SComment := TempFileName(ExtractFilePath(SComment)+ExtractFileNameWithoutExt(SComment), 'txt');
          If Not ForceDirectories(ExtractFilePath(SComment)) then begin ErrMsg('������ �������� ��������: '+ExtractFilePath(SComment)); Exit; end;
          CBComment.Lines.SaveToFile(SComment);
       end;

       {��������� ������}
       //  t   - �������������� �����
       //  ap  - ������ ���� ������ ������
       //  ep1 - ��������� �� ����� ������� �������
       //  dh  - ��������� ��������� ������������ �����
       //  hp  - ������
       //  k   - ������������� �����
       //  m   - ����� ������
       SKey := 'a -t -dh -m5 -ep1 -ap'+CBYear.Text;
       If PasRar  <> '' then SKey := SKey + ' -hp"'+PasRar+'"';
       If SComment = '' then SKey := SKey + ' -k';
(*
       {�������}
       SRar:=ReplStr(SRar, '7zFM.exe', '7z.exe');
       SRar:=ReplStr(SRar, '7zG.exe',  '7z.exe');

       SKey := 'a -ap'+CBYear.Text;
       If PasRar <> '' then SKey := SKey + ' -p"'+PasRar+'"';
       {������ �����������}
       If Not ExecAndWait(SRar, SKey+' '+'"'+SDect+'" "'+PATH_BD_DATA+CBYear.Text+'\'+CBMonth.Text+'"', SW_SHOWNORMAL) then begin
          ErrMsg('������ �������� ����� ��������!');
          If FileExists(SDect) then DeleteFile(SDect);
          Exit;
       end;
*)
       {������ �����������}
       If Not ExecAndWait(SRar, SKey+' -- '+'"'+SDect+'" "'+PATH_BD_DATA+CBYear.Text+'\'+CBMonth.Text+'"', SW_SHOWNORMAL) then begin
          ErrMsg('������ �������� ����� ��������!');
          If FileExists(SDect) then DeleteFile(SDect);
          Exit;
       end;

       {������ �����������}
       If SComment<>'' then begin
          SKey := 'c -m5 -k -z"'+SComment+'"';
          If PasRar  <> '' then SKey := SKey + ' -hp"'+PasRar+'"';
          ExecAndWait(SRar, SKey+' -- '+'"'+SDect+'"', SW_MINIMIZE);
       end;

       {���� ����� ���������� ���, �� ������}
       If FileExists(SDect) then begin
          ForegroundWindow;
          MessageDlg('���� �������� ������!'+CH_NEW+SDect, mtInformation, [mbOk], 0);
          StartAssociatedExe('"'+ExtractFilePath(SDect)+'"', SW_SHOWNORMAL);
       end else begin
          ErrMsg('������ �������� ����� ��������!');
          Exit;
       end;

    finally
       {����������}
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
       {������� ���� �����������}
       If SComment<>'' then If FileExists(SComment) then DeleteFile(SComment);
       {������������ ���������}
       ModalResult := mrNo;
    end;

    ModalResult := mrOk;
end;


{==============================================================================}
{============================   ACTION: ������   ==============================}
{==============================================================================}
procedure TFEXPORT.ACancelExecute(Sender: TObject);
begin
    ModalResult := mrCancel;
end;


{==============================================================================}
{========================   �������: ��������� ����   =========================}
{==============================================================================}
procedure TFEXPORT.CBYearChange(Sender: TObject);
begin
    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{======================   �������: ��������� ������   =========================}
{==============================================================================}
procedure TFEXPORT.CBMonthChange(Sender: TObject);
begin
    {����������� Action}
    EnablAction;
end;

{==============================================================================}
{====================   �������: ��������� ����������   =======================}
{==============================================================================}
procedure TFEXPORT.CBFolderChange(Sender: TObject);
begin
    {���������� ���� �������� ������}
    WriteLocalString(INI_SET, INI_SET_EXPORT_PATH, CBFolder.Text);
    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{======================   �������: ����� ����������   =========================}
{==============================================================================}
procedure TFEXPORT.CBFolderRightButtonClick(Sender: TObject);
var SPath : String;
begin
    SPath := CBFolder.Text;
    If SelectDirectory('������� �������� ������', '', SPath) then begin
       If SPath[Length(SPath)] <> '\' then SPath:=SPath+'\';
       CBFolder.Text := SPath;
    end;
end;

{==============================================================================}
{=========================   ��� �� ������� ����   ============================}
{==============================================================================}
function TFEXPORT.IsExistsAllRegions: Boolean;
const SEPARATOR = '; ';
var LReg : TStringList;
    IReg : Integer;
    SErr : String;
begin
    {�������������}
    Result := false;
    SErr   := '';
    LReg   := TStringList.Create;
    try
       {��������� ������� ������}
       If Not FFMAIN.ListRegionsRangeDown(@LReg, SMAIN_REGION, CBYear.Text, CBMonth.Text) then Exit;
       For IReg := 0 to LReg.Count -1 do begin
           If Not FileExists(FFMAIN.PathRegion(CBYear.Text, CBMonth.Text, LReg[IReg])) then begin
              SErr := SErr + LReg[IReg] + SEPARATOR;
              If Length(SErr) > 200 then begin
                 Delete(SErr, 200, Length(SErr));
                 SErr := SErr + '...'+SEPARATOR;
                 Break;
              end;
           end;
       end;
    finally
       LReg.Free;
    end;
    If SErr <> '' then Delete(SErr, Length(SErr) - Length(SEPARATOR) + 1, Length(SEPARATOR));

    {�������������}
    If SErr <> '' then begin
       If MessageDlg('����������� ��������� ������:'+CH_NEW+SErr+CH_NEW+CH_NEW+
                     '����������� ��������.', mtWarning, [mbOk, mbCancel], 0) <> mrOk then Exit;
    end;

    {���������}
    Result := true;
end;

end.
