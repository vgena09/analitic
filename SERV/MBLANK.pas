unit MBLANK;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.FileCtrl,
  Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.CheckLst, Vcl.ActnList, Vcl.Buttons,
  IdGlobal, IdGlobalProtocols,
  MAIN;

type
  TFBLANK = class(TForm)
    CBView: TCheckListBox;
    LCaption: TLabel;
    PSpace1: TPanel;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel3: TPanel;
    BtnCancel: TBitBtn;
    AList: TActionList;
    AOk: TAction;
    ACancel: TAction;
    Panel4: TPanel;
    BtnOk: TBitBtn;
    Panel5: TPanel;
    LFolder: TLabel;
    CBFolder: TButtonedEdit;
    procedure FormCreate(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure CBFolderChange(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure CBViewClick(Sender: TObject);
    procedure CBViewKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure AOkExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure CBFolderRightButtonClick(Sender: TObject);
  private
    FFMAIN    : TFMAIN;
    DatPeriod : TDate;

    procedure EnablAction;
  public

  end;

var
  FBLANK: TFBLANK;

implementation

uses FunConst, FunText, FunSys, FunVcl, FunBD, FunDay, FunExcel, FunIni;

{$R *.dfm}


{==============================================================================}
{============================   �������� �����    =============================}
{==============================================================================}
procedure TFBLANK.FormCreate(Sender: TObject);
var LForma_ : TLForma;
    I       : Integer;
begin
    {�������������}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ������ ������ ������ ��� ��������';
    DatPeriod := DateTimeCorrect(Now, 0, 1, 0, 0, 0);

    {��������� ������ �������}
    CBView.Items.Clear;
    try
       If Not FFMAIN.FindForms(@LForma_, false, true, -1, -1, '', '', '', '', false, DatPeriod) then begin
          ErrMsg('������ ������������ ������ �������!');
          Exit;
       end;
       For I:=Low(LForma_) to High(LForma_) do begin
          If Not LForma_[I].FExport   then Continue;
          If LForma_[I].FCaption = '' then Continue;
          CBView.Items.Add(LForma_[I].FCaption);
       end;
    finally
       SetLength(LForma_, 0);
    end;
    {������ ���� �������� ������ �� ����� ������� ��������}
    CBFolder.Text := ReadLocalString(INI_SET, INI_SET_EXPORT_PATH, '');

    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{============================   ��������� �����    ============================}
{==============================================================================}
procedure TFBLANK.FormActivate(Sender: TObject);
begin
    CBView.SetFocus; 
end;


{==============================================================================}
{=============================   ������� �����    =============================}
{==============================================================================}
procedure TFBLANK.FormHide(Sender: TObject);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;


{==============================================================================}
{========================   ACTION: ������ �������    =========================}
{==============================================================================}
procedure TFBLANK.AOkExecute(Sender: TObject);
var LForma_ : TLForma;
    SSrc, STemp, SDect: String;
    I : Integer;
    E : Variant;
begin
    {�������������}
    ModalResult:=mrCancel;

    {���� ���������� ������}
    For I:=0 to CBView.Count-1 do begin
       If Not CBView.Checked[I] then Continue;
       FFMAIN.Repaint; 

       try {���� ���������� �����}
           If Not FFMAIN.FindForms(@LForma_, true, true, -1, -1, '', CBView.Items[I], '', '', false, DatPeriod) then begin ErrMsg('�� ������� ������ ������: '+CBView.Items[I]); Continue; end;
           {����-��������}
           SSrc:=FFMAIN.PathFileForm(LForma_[0].SBankPref, LForma_[0].FFile);
           If Not FileExists(SSrc) then begin ErrMsg('�� ������ ���� ������: '+SSrc); Continue; end;
       finally
           SetLength(LForma_, 0);
       end;

       {����-���������}
       SDect:=CBFolder.Text;
       If SDect[Length(SDect)] <> '\' then SDect:=SDect+'\';
       SDect:=SDect+ExtractFileName(SSrc);

       {������� ���������� ���� ����������}
       If FileExists(SDect) then begin
          If MessageDlg('���� � ����������� ������ ��� ����������:'+CH_NEW+SDect+CH_NEW+'������� ������ ����?',
                        mtWarning, [mbYes, mbNo], 0)<>mrYes then begin ModalResult:=mrNone; Continue; end;
          FFMAIN.Repaint; 
          DeleteFile(SDect);
       end;

       {����������}
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ���������� ������ ������: '+CBView.Items[I];
       FFMAIN.Repaint; 

       {��������� ����}
       STemp := CopyFileToTemp(SSrc);
       If STemp='' then begin ErrMsg('������ �������� ���������� �����!'); Continue; end;

       {������������ �����}
       If Not CreateExcel then begin ErrMsg('������ �������� MS Excel!'); Continue; end;
       try If Not OpenXls(STemp, true) then begin ErrMsg('������ �������� �����: '+CH_NEW+STemp); Continue; end;
           try
              E:=GetExcelApplicationID;
              E.DisplayAlerts:=false;
              E.ActiveWorkbook.WorkSheets[1].Range[CELL_FULL].Value:='FULL';
              If Not SaveXls then begin ErrMsg('������ ���������� ���������� �����: '+CH_NEW+STemp); Continue; end;
           finally
              If Not CloseXls then ErrMsg('������ �������� �������������� ������!');
           end;
       finally
          If Not CloseExcel then ErrMsg('������ �������� Excel!');
          E := Unassigned;
       end;

       {�������� ����}
       If Not ForceDirectories(ExtractFilePath(SDect)) then begin ErrMsg('������ �������� ��������: '+ExtractFilePath(SDect)); Continue; end;
       If Not CopyFileTo(STemp, SDect)                 then begin ErrMsg('������ �������� ����� ������: '+SDect); Continue; end;

       {������� ��������� ����}
       DeleteFile(STemp);
    end;

    {��������� �����}
    StartAssociatedExe('"'+ExtractFilePath(SDect)+'"', SW_SHOWNORMAL);

    {������������ ���������}
    ModalResult:=mrOk;
end;



{==============================================================================}
{===========================   ACTION: ������    ==============================}
{==============================================================================}
procedure TFBLANK.ACancelExecute(Sender: TObject);
begin
    ModalResult := mrCancel;
end;


{==============================================================================}
{==========================   ����������� ACTION   ============================}
{==============================================================================}
procedure TFBLANK.EnablAction;
var I : Integer;
    B : Boolean;
begin
    {������ ���� ���� ���� �����}
    B := false;
    For I:=0 to CBView.Count-1 do begin
       If CBView.Checked[I] then B:=true;
    end;

    {������� ������ ���� �����}
    If Not DirectoryExists(CBFolder.Text) then B:=false;

    {������������� �����������}
    AOk.Enabled := B;
end;


{==============================================================================}
{====================   �������: ��������� ����������   =======================}
{==============================================================================}
procedure TFBLANK.CBFolderChange(Sender: TObject);
begin
    {���������� ���� �������� ������}
    WriteLocalString(INI_SET, INI_SET_EXPORT_PATH, CBFolder.Text);
    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{======================   �������: ����� ����������   =========================}
{==============================================================================}
procedure TFBLANK.CBFolderRightButtonClick(Sender: TObject);
var SPath : String;
begin
    SPath := CBFolder.Text;
    If SelectDirectory('����� ���������� �������', '', SPath) then begin
       If SPath[Length(SPath)] <> '\' then SPath:=SPath+'\';
       CBFolder.Text := SPath;
    end;
end;

{==============================================================================}
{==========================   �������: ON_CLICK   =============================}
{==============================================================================}
procedure TFBLANK.CBViewClick(Sender: TObject);
begin
    {����������� Action}
    EnablAction;
end;

procedure TFBLANK.CBViewKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    CBViewClick(Sender);
end;

end.
