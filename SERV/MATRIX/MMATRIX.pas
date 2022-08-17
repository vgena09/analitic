unit MMATRIX;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.OleServer,
  MAIN,
  ExcelXP, Excel2000,
  ElEdits, ElBtnEdit, ElSpin, ElCheckCtl, ElPopBtn, ElToolBar, ActnList,
  ElXPThemedControl, ElBtnCtl, ExtCtrls, ElPanel,
  LMDControl, LMDBaseControl, LMDBaseGraphicControl, LMDBaseLabel,
  LMDCustomSimpleLabel, LMDSimpleLabel;

type
  TFMATRIX = class(TForm)
    PBottom: TElPanel;
    ElPopupButton1: TElPopupButton;
    ElPanel5: TElPanel;
    XLS: TExcelApplication;
    WB: TExcelWorkbook;
    ActionList1: TActionList;
    ARun: TAction;
    AClear: TAction;
    ElPanel1: TElPanel;
    AOpen: TAction;
    ElPanel14: TElPanel;
    ElPanel10: TElPanel;
    PBlockKod: TElPanel;
    LBlockKod: TLMDSimpleLabel;
    ElPanel2: TElPanel;
    PPage: TElPanel;
    LPage: TLMDSimpleLabel;
    EPage: TElSpinEdit;
    ElPanel3: TElPanel;
    ElPanel4: TElPanel;
    PBlockIn: TElPanel;
    LBlockIn: TLMDSimpleLabel;
    EBlockIn: TElButtonEdit;
    ElPanel11: TElPanel;
    PBlockOut: TElPanel;
    LBlockOut: TLMDSimpleLabel;
    EBlockOut: TElButtonEdit;
    EBlockKod: TElButtonEdit;
    ElPopupButton2: TElPopupButton;
    ElPopupButton3: TElPopupButton;
    ElPopupButton4: TElPopupButton;
    procedure FormCreate(Sender: TObject);
    procedure EChange(Sender: TObject);
    procedure ARunExecute(Sender: TObject);
    procedure AClearExecute(Sender: TObject);
    procedure EBlockInEnter(Sender: TObject);
    procedure EBlockInExit(Sender: TObject);
    procedure EBlockOutEnter(Sender: TObject);
    procedure EBlockOutExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure EBlockOutButtonClick(Sender: TObject);
    procedure EBlockInButtonClick(Sender: TObject);
    procedure EBlockOutAltButtonClick(Sender: TObject);
    procedure AOpenExecute(Sender: TObject);
    procedure EBlockKodButtonClick(Sender: TObject);
    procedure EBlockInAltButtonClick(Sender: TObject);
    procedure WBBeforeClose(ASender: TObject; var Cancel: WordBool);
    procedure WBSheetSelectionChange(ASender: TObject; const Sh: IDispatch;
      const Target: ExcelRange);
  private
    FFMAIN             : TFMAIN;
    SPathXLS, SPathINI : String;
    IsIn, IsOut        : Boolean;     // Фокус ввода
    SOut               : String;      // Неизменяемая строка

    procedure EnablAction;
    function  ConnectExcel(const SPath: String): Boolean;
    procedure DisconnectExcel(const IsClose: Boolean);
  public
    procedure Execute(const SPath: String);
  end;

var
  FMATRIX: TFMATRIX;

implementation

uses FunConst, FunText, FunExcel, FunSys, FunVcl, FunBD,
     MKOD;

{$R *.dfm}
{$INCLUDE MMATRIX_ACTION}
{$INCLUDE MMATRIX_EXCEL}


{==============================================================================}
{===========================   ВНЕШНИЙ ВЫЗОВ    ===============================}
{==============================================================================}
procedure TFMATRIX.Execute(const SPath: String);
begin
    {Инициализация}
    If Not FileExists(SPath) then Exit;
    SPathXLS := SPath;
    SPathINI := ChangeFileExt(SPathXLS, '.ini');
    If Not FileExists(SPathINI) then Exit;
    Caption  := 'Мастер создания матрицы: '+ExtractFileNameWithoutExt(SPath);
    EChange(nil);

    {Перерисовка окон перед длительной операцией}
    FFMAIN.Repaint;
    Repaint;

    {Открытие EXCEL}
    ConnectExcel(SPath);

    {Окно программы на передний план}
    ForegroundWindow;

    {Режим диалога}
    ShowModal;
end;


{==============================================================================}
{==========================   СОЗДАНИЕ ФОРМЫ    ===============================}
{==============================================================================}
procedure TFMATRIX.FormCreate(Sender: TObject);
begin
    {Инициализация}
    FFMAIN             := TFMAIN(GlFindComponent('FMAIN'));
    IsIn               := false;
    IsOut              := false;
    SOut               := '';

    WB.OnSheetSelectionChange := WBSheetSelectionChange;
    WB.OnBeforeClose          := WBBeforeClose;

    EBlockIn.ButtonIsSwitch   := true;
    EBlockIn.ButtonDown       := false;

    EBlockKod.OnChange  := EChange;
    EBlockIn.OnChange   := EChange;
    EBlockOut.OnChange  := EChange;

    EBlockIn.ButtonDown := true;
    EBlockIn.OnEnter    := EBlockInEnter;
    EBlockIn.OnExit     := EBlockInExit;
    EBlockOut.OnEnter   := EBlockOutEnter;
    EBlockOut.OnExit    := EBlockOutExit;
end;


{==============================================================================}
{==========================   ЗАКРЫТИЕ ФОРМЫ    ===============================}
{==============================================================================}
procedure TFMATRIX.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {Отключение от EXCEL}
    DisconnectExcel(true);
end;


{==============================================================================}
{==============================   ВЫБОР КОДА   ================================}
{==============================================================================}
procedure TFMATRIX.EBlockKodButtonClick(Sender: TObject);
var F: TFKOD;
begin
    F:=TFKOD.Create(Self);
    try     EBlockKod.Text := F.Execute(EBlockKod.Text, false);
    finally F.Free; end;
    
    {Доступность Action}
    EnablAction;
end;


{==============================================================================}
{=========================   ПРИ ИЗМЕНЕНИИ ПОЛЕЙ   ============================}
{==============================================================================}
procedure TFMATRIX.EChange(Sender: TObject);
begin
    {При синхронизации поле вывода недоступно}
    EBlockOut.Enabled          := Not EBlockIn.ButtonDown;
    EBlockOut.ButtonEnabled    := Not EBlockIn.ButtonDown;
    EBlockOut.AltButtonEnabled := Not EBlockIn.ButtonDown;
    EBlockIn.AltButtonEnabled  := Not EBlockIn.ButtonDown;
    If EBlockIn.ButtonDown then EBlockOut.Font.Color := clGrayText
                           else EBlockOut.Font.Color := clWindowText;

    {Доступность Action}
    EnablAction;
end;

procedure TFMATRIX.EBlockOutAltButtonClick(Sender: TObject);
begin SOut:=EBlockOut.Text; end;
procedure TFMATRIX.EBlockInEnter(Sender: TObject);
begin IsIn:=true;  end;
procedure TFMATRIX.EBlockInExit(Sender: TObject);
begin IsIn:=false; end;
procedure TFMATRIX.EBlockOutEnter(Sender: TObject);
begin IsOut:=true;  SOut:=EBlockOut.Text; end;
procedure TFMATRIX.EBlockOutExit(Sender: TObject);
begin IsOut:=false; SOut:=''; end;

end.
