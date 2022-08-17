unit MBDREGION;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.FileCtrl,
  Vcl.Buttons, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ActnList,
  IdGlobalProtocols,
  MAIN;

type
  TFBDREGION = class(TForm)
    ActionList1: TActionList;
    AOK: TAction;
    ACancel: TAction;
    PRegion: TPanel;
    LRegion: TLabel;
    PFolder: TPanel;
    LFolder: TLabel;
    CBFolder: TButtonedEdit;
    CBRegion: TComboBox;
    Panel3: TPanel;
    BitCancel: TBitBtn;
    BtnOk: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CBRegionChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure AOKExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
    procedure CBFolderChange(Sender: TObject);
    procedure CBFolderRightButtonClick(Sender: TObject);
  private
    FFMAIN : TFMAIN;
    procedure EnablAction;
  public
  
  end;

var
  FBDREGION : TFBDREGION;

implementation

uses FunConst, FunSys, FunVcl, FunBD, FunText, FunDay, FunIni;

{$R *.dfm}

{==============================================================================}
{==========================   �������� �����   ================================}
{==============================================================================}
procedure TFBDREGION.FormCreate(Sender: TObject);
var I: Integer;
begin                                                  
    {�������������}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text := ' ������� ������ ��� ��������';

    {��������� ������ ��������}
    CBRegion.AutoComplete      := true;
    CBRegion.Style             := csDropDown;
    CBRegion.OnChange          := CBRegionChange; // nil
    CBRegion.OnSelect          := CBRegionChange;
    CBRegion.OnKeyDown         := FormKeyDown;
    CBRegion.Items.Clear;
    For I:=Low(FFMAIN.LRegion) to High(FFMAIN.LRegion) do CBRegion.Items.Add(FFMAIN.LRegion[I].FCaption);

    {������ ���� �������� ������ �� ����� ������� ��������}
    CBFolder.Text := ReadLocalString(INI_SET, INI_SET_EXPORT_PATH, '');

    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{==========================   �������� �����   ================================}
{==============================================================================}
procedure TFBDREGION.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;


{==============================================================================}
{=========================   ACTION: ���������   ==============================}
{==============================================================================}
procedure TFBDREGION.AOKExecute(Sender: TObject);
label NxYear, NxMonth, NxRegion;
var FDYear, FDMonth, FDRegion : TWIN32FindData;
    FHYear, FHMonth, FHRegion : THandle;
    LReg                      : TStringList;
    SPath, SPathDest          : String;
    SYear, SMonth, SRegion    : String;
    SRegion1, SRegionMain     : String;

begin
    {�������������}
    Hide;
    FFMAIN.Repaint;
    SRegionMain := CBRegion.Text;
    SPathDest   := CBFolder.Text;
    FHRegion    := 0;
    FHMonth     := 0;

    LReg:=TStringList.Create;
    try
       {������ �������� �� ����������}
       If Not FFMAIN.ListRegionsRangeDown(@LReg, SRegionMain, '', '') then Exit;

       {����� �� �����}
       FHYear := FindFirstFile(PChar(PATH_BD_DATA+'*'), FDYear);
       While FHYear <> INVALID_HANDLE_VALUE do begin
          SYear:=String(FDYear.cFileName);
          If Not IsIntegerStr(SYear) then Goto NxYear;

          {����� �� �������}
          FHMonth := FindFirstFile(PChar(PATH_BD_DATA+SYear+'\*'), FDMonth);
          While FHMonth <> INVALID_HANDLE_VALUE do begin
             SMonth:=String(FDMonth.cFileName);
             If MonthStrToInd(SMonth)<Low(MonthList) then Goto NxMonth;

             {����������}
             FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' ��������: '+SMonth+' '+SYear+' �.';
             FFMAIN.Refresh;
             Application.ProcessMessages;
                                            
             {����� �� ��������}
             FHRegion := FindFirstFile(PChar(PATH_BD_DATA+SYear+'\'+SMonth+'\*.dat'), FDRegion);
             While FHRegion <> INVALID_HANDLE_VALUE do begin
                SRegion  := String(FDRegion.cFileName);
                SRegion1 := ReplModulStr(SRegion, '', SUB_RGN1, SUB_RGN2);
                Delete(SRegion1, Length(SRegion1)-4+1, 4);

                {�������� ������}
                If LReg.IndexOf(SRegion1)<0 then Goto NxRegion;
                ForceDirectories(SPathDest+'\'+SYear+'\'+SMonth);
                SPath:=SYear+'\'+SMonth+'\'+SRegion;
                If Not CopyFileTo(PATH_BD_DATA+SPath, SPathDest+'\'+SPath) then ErrMsg('������ �����������!');

NxRegion:       {��������� ������}
                If Not FindNextFile(FHRegion, FDRegion) then Break;
             end;
             Winapi.Windows.FindClose(FHRegion);
NxMonth:     {��������� �����}
             If Not FindNextFile(FHMonth, FDMonth) then Break;
          end;
          Winapi.Windows.FindClose(FHMonth);
NxYear:   {��������� ���}                                 
          If Not FindNextFile(FHYear, FDYear) then Break;
       end;
       Winapi.Windows.FindClose(FHYear);

       {�����������}
       MessageDlg('���� ������ �����������!'+CH_NEW+SPathDest, mtInformation, [mbOk], 0);
       StartAssociatedExe('"'+SPathDest+'"', SW_SHOWNORMAL);

    finally
       FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
       {������������ ���������}
       ModalResult:=mrOk;
    end;
end;


{==============================================================================}
{===========================   ACTION: ������   ===============================}
{==============================================================================}
procedure TFBDREGION.ACancelExecute(Sender: TObject);
begin
    ModalResult:=mrCancel;
end;


{==============================================================================}
{========================   ����������� ACTION   ==============================}
{==============================================================================}
procedure TFBDREGION.EnablAction;
var B_OK: Boolean;
begin
    {�������������}
    B_OK:=true;

    {�������� �������}
    If CBRegion.Items.IndexOf(CBRegion.Text)>=0 then CBRegion.Font.Color := clBlack
                                                else CBRegion.Font.Color := clRed;
    {�������� ��������}
    If DirectoryExists(CBFolder.Text) then CBFolder.Font.Color := clBlack
                                      else CBFolder.Font.Color := clRed;

    {��������}
    If (CBRegion.Font.Color=clRed) or (CBFolder.Font.Color=clRed) then B_OK:=false;

    {�����������}
    AOK.Enabled := B_OK;
end;


{==============================================================================}
{========================   ��������� �������   ===============================}
{==============================================================================}
procedure TFBDREGION.CBRegionChange(Sender: TObject);
begin
    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{============================   ������� �������   =============================}
{==============================================================================}
procedure TFBDREGION.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    Case Key of
    13: {ENTER} If AOK.Enabled then AOK.Execute;
    27: {ESC}   ModalResult:=mrCancel;
    end;
end;


{==============================================================================}
{====================   �������: ��������� ����������   =======================}
{==============================================================================}
procedure TFBDREGION.CBFolderChange(Sender: TObject);
begin
    {���������� ���� �������� ������}
    WriteLocalString(INI_SET, INI_SET_EXPORT_PATH, CBFolder.Text);
    {����������� Action}
    EnablAction;
end;


{==============================================================================}
{======================   �������: ����� ����������   =========================}
{==============================================================================}
procedure TFBDREGION.CBFolderRightButtonClick(Sender: TObject);
var SPath : String;
begin
    SPath := CBFolder.Text;
    If SelectDirectory('����� ���������� �� �������', '', SPath) then begin
       If SPath[Length(SPath)] <> '\' then SPath:=SPath+'\';
       CBFolder.Text := SPath;
    end;
end;

end.
