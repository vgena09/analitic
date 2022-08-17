unit MCLOSE;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ActnList,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons, Vcl.Imaging.jpeg,
  MAIN;

type
  TFCLOSE = class(TForm)
    ActionList1: TActionList;
    AExit: TAction;
    ACancel: TAction;
    CBArj: TCheckBox;
    Image1: TImage;
    PBottom: TPanel;
    BtnCancel: TBitBtn;
    BtnOk: TBitBtn;
    Bevel1: TBevel;
    procedure AArjExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure AExitExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
  private
    FFMAIN           : TFMAIN;
    SPATH_ARJ, SDect : String;
  public
    function  Execute: Boolean;
  end;

const
  FILE_ARJ = '�����';

var
  FCLOSE: TFCLOSE;

implementation

uses FunConst, FunSys, FunInfo, FunVcl, FunFiles, FunDay, FunText, FunIni;

{$R *.dfm}


{==============================================================================}
{============================   ������� �����   ===============================}
{==============================================================================}
function TFCLOSE.Execute: Boolean;
begin
    Result := (ShowModal=mrOk);
    If Result and CBArj.Checked then AArjExecute(nil);
end;


{==============================================================================}
{============================   �������� �����   ==============================}
{==============================================================================}
procedure TFCLOSE.FormCreate(Sender: TObject);
var Sr                        : TSearchRec;
    CurDate, LastDate         : TDateTime;
    Hours, Days, Month, Years : Integer;
begin
    {�������������}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));

    {��� ��������������}
    If IS_ADMIN then begin
       {���� � ����� �������}
       SPATH_ARJ := ReadLocalString(INI_SET, INI_SET_ARJ_PATH, PATH_PROG+PATH_ARJ);
       If Length(SPATH_ARJ) > 0 then begin
          If SPATH_ARJ[Length(SPATH_ARJ)] <> '\' then SPATH_ARJ:=SPATH_ARJ+'\';
       end;

       {���� � ����� ��������� ���������}
       LastDate:=0;
       try If FindFirst(SPATH_ARJ+FILE_ARJ+'*.rar', faDirectory, Sr) = 0 then begin
              repeat
                 CurDate := GetFileDate(SPATH_ARJ+Sr.Name);
                 If CurDate > LastDate then LastDate:=CurDate;
              until FindNext(Sr) <> 0;
           end;
       finally FindClose(Sr);
       end;
       DateTimeDiff0(LastDate, Now, Hours, Days, Month, Years);

    {��� ������������}
    end else begin
       CBArj.Visible := false;
       PBottom.Top   := PBottom.Top - CBArj.Height * 2;
       Height        := Height - CBArj.Height * 2;
       Years         := 0;
       Month         := 0;
    end;

    {���� ������ ����� ������}
    CBArj.Checked := IS_NET and ((Years > 0) or (Month > 0));
end;


{==============================================================================}
{======================   ACTION: ����� �� ���������   ========================}
{==============================================================================}
procedure TFCLOSE.AExitExecute(Sender: TObject);
begin
    ModalResult := mrOk;
end;


{==============================================================================}
{===========================   ACTION: ������   ===============================}
{==============================================================================}
procedure TFCLOSE.ACancelExecute(Sender: TObject);
begin
    ModalResult := mrCancel;
end;


{==============================================================================}
{===========================   ACTION: ���������   ============================}
{==============================================================================}
procedure TFCLOSE.AArjExecute(Sender: TObject);
var SRar : String;
    SSrc : String;
begin
    try
       {����������� �����}
       Hide;
       FFMAIN.Repaint;

       {���� ����������}
       SRar := GetProgramAssociation('.rar');
       If Not FileExists(SRar) then begin ErrMsg('�� ���������� ��������� ��� ������ ������ � RAR-�����!'); Exit; end;

       {���� ������}
       ForceDirectories(SPATH_ARJ);
       SDect:=TempFileName(SPATH_ARJ+FILE_ARJ, 'rar');

       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ��������� ������������� ���� ������';
       {�������������}
       //  t   - �������������� �����
       //  dh  - ��������� ��������� ������������ �����
       //  hp  - ������
       //  k   - ������������� �����
       //  m   - ����� ������
       SSrc:=PATH_BD_DATA;
       TokCharEnd(SSrc, '\');
       If Not ExecAndWait(SRar, 'a -t -dh -m5 -k -- '+'"'+SDect+'" "'+SSrc+'"', SW_SHOWNORMAL) then begin
          ErrMsg('������ �������� ����� ������!');
          If FileExists(SDect) then DeleteFile(SDect);
          Exit;
       end;
    finally
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
       ModalResult:=mrOk;
    end;
end;

end.
