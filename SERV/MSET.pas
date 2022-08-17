unit MSET;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.FileCtrl,
  Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Data.DB, Data.Win.ADODB,
  MAIN;

type
  TFSET = class(TForm)
    PBottom: TPanel;
    BtnClose: TBitBtn;
    BtnReset: TBitBtn;
    Panel1: TPanel;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TSAny: TTabSheet;
    CBModifyKod: TCheckBox;
    CBImportSubReg: TCheckBox;
    PPathArj: TPanel;
    CBVerifyRAR: TCheckBox;
    CBVerifyXLS: TCheckBox;
    Bevel3: TBevel;
    LPathArj: TLabel;
    PPathExport: TPanel;
    LPathExport: TLabel;
    Bevel4: TBevel;
    PPas: TPanel;
    LPas: TLabel;
    CBPas: TEdit;
    PMainReg: TPanel;
    LMainReg: TLabel;
    EMainReg: TComboBox;
    CBNet: TCheckBox;
    CBTerminate: TCheckBox;
    Bevel2: TBevel;
    GroupBox1: TGroupBox;
    CBSEPDetail: TCheckBox;
    Panel3: TPanel;
    LSEPMonth: TLabel;
    EPathArj: TButtonedEdit;
    EPathExport: TButtonedEdit;
    TBSEPMonth: TTrackBar;
    LSEPMonth2: TLabel;
    procedure FormShow(Sender: TObject);
    procedure FormHide(Sender: TObject);
    procedure ElPopupButton1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure CBImportSubRegClick(Sender: TObject);
    procedure CBImportSubRegKeyPress(Sender: TObject; var Key: Char);
    procedure CBVerifyRARClick(Sender: TObject);
    procedure CBVerifyRARKeyPress(Sender: TObject; var Key: Char);
    procedure CBVerifyXLSClick(Sender: TObject);
    procedure CBVerifyXLSKeyPress(Sender: TObject; var Key: Char);
    procedure CBModifyKodClick(Sender: TObject);
    procedure CBModifyKodKeyPress(Sender: TObject; var Key: Char);
    procedure EPathExportChange(Sender: TObject);
    procedure EPathArjChange(Sender: TObject);
    procedure CBSEPDetailClick(Sender: TObject);
    procedure CBSEPDetailKeyPress(Sender: TObject; var Key: Char);
    procedure TBSEPMonthChange(Sender: TObject);
    procedure CBTerminateClick(Sender: TObject);
    procedure CBTerminateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EMainRegChange(Sender: TObject);
    procedure CBNetClick(Sender: TObject);
    procedure CBNetKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure EPathArjRightButtonClick(Sender: TObject);
    procedure EPathExportRightButtonClick(Sender: TObject);
  private
    FFMAIN   : TFMAIN;
    Q_REGION : TADOQuery;
    procedure EnablSEPMonth;
  public

  end;

var
  FSET   : TFSET;

implementation

uses FunConst, FunSys, FunVcl, FunText, FunBD, FunIni;

{$R *.dfm}


{==============================================================================}
{=============================   СОЗДАНИЕ ФОРМЫ    ============================}
{==============================================================================}
procedure TFSET.FormCreate(Sender: TObject);
begin
    FFMAIN:= TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Настройка программы';
end;


{==============================================================================}
{===============================   ПОКАЗ ФОРМЫ    =============================}
{==============================================================================}
procedure TFSET.FormShow(Sender: TObject);
begin
    {Главный регион}
    EMainReg.AutoComplete := true;
    EMainReg.Style          := csDropDownList;
    Q_REGION := TADOQuery.Create(Self);
    try
       Q_REGION.Connection := FFMAIN.LBank[0].BD;
       Q_REGION.SQL.Text := 'SELECT ['+F_COUNTER    +'], ['+F_CAPTION+']'+CH_NEW+
                            'FROM   ['+TABLE_REGIONS+']'+CH_NEW+
                            'WHERE  ((:Период_начало < ['+F_END  +']) OR (['+F_END  +'] Is Null)) AND '+         // не работает условие: OR (:Период_начало Is Null)
                                   '((:Период_конец  > ['+F_BEGIN+']) OR (['+F_BEGIN+'] Is Null))'+CH_NEW+
                            'ORDER BY ['+F_CAPTION+'] ASC;';
       // Q_REGION.Parameters.ParseSQL(Q_REGION.SQL.Text, True);
       Q_REGION.Parameters.ParamByName('Период_начало').Value := Date;
       Q_REGION.Parameters.ParamByName('Период_конец').Value  := Date;
       Q_REGION.Open;
       Q_REGION.First;
       While Not Q_REGION.Eof do begin
          EMainReg.Items.Add(Q_REGION.FieldByName(F_CAPTION).AsString);
          If IMAIN_REGION=Q_REGION.FieldByName(F_COUNTER).AsInteger then EMainReg.ItemIndex := EMainReg.Items.Count-1;
          Q_REGION.Next;
       end;
       //Q_REGION.Close;
    finally
       //Q_REGION.Free;
    end;

    {Признак работы в сети}
    CBNet.Checked := IS_NET;

    {Признак отключения пользователей}
    CBTerminate.Checked := IS_TERMINATE;

    {Читаем пароль}
    CBPas.Text := ReadGlobalString(INI_SET, INI_SET_PAS_IMPORT_EXPORT, '');

    {Читаем локальные настройки}
    CBImportSubReg.Checked := ReadLocalBool   (INI_SET, INI_SET_IMPORT_SUBREG,     true);
    CBVerifyRAR.Checked    := ReadLocalBool   (INI_SET, INI_SET_IMPORT_VERIFY_RAR, true);
    CBVerifyXLS.Checked    := ReadLocalBool   (INI_SET, INI_SET_IMPORT_VERIFY_XLS, true);
    EPathArj.Text          := ReadLocalString (INI_SET, INI_SET_ARJ_PATH,          PATH_PROG+PATH_ARJ);
    EPathExport.Text       := ReadLocalString (INI_SET, INI_SET_EXPORT_PATH,       '');

    CBModifyKod.Checked    := ReadLocalBool   (INI_SET, INI_SET_KOD_MODIFY,        false);

    CBSEPDetail.OnClick    := nil;
    CBSEPDetail.Checked    := ReadLocalBool   (INI_SET, INI_SET_SEP_DETAIL,        false);
    CBSEPDetail.OnClick    := CBSEPDetailClick;
    CBSEPDetail.OnKeyPress := CBSEPDetailKeyPress;
    EnablSEPMonth;

    TBSEPMonth.OnChange    := nil;
    TBSEPMonth.Min         := 12;
    TBSEPMonth.Max         := 36;
    TBSEPMonth.Position    := ReadLocalInteger(INI_SET, INI_SET_SEP_MONTH,         18);
    TBSEPMonth.OnChange    := TBSEPMonthChange;
    TBSEPMonthChange(nil);

    {Недоступность некоторых элементов без прав администратора}
    If Not IS_ADMIN then begin
       LMainReg.Enabled  := false or (IMAIN_REGION = 0);
       EMainReg.Enabled  := false or (IMAIN_REGION = 0);
       CBNet.Enabled     := false;

       CBPas.Enabled     := false;
       LPas.Enabled      := false;
    end;

    {Недоступность некоторых элементов при работе в сети}
    If IS_NET then begin
       CBModifyKod.Enabled := false;
    end;
end;


{==============================================================================}
{==============================   СКРЫТИЕ ФОРМЫ    ============================}
{==============================================================================}
procedure TFSET.FormHide(Sender: TObject);
begin
    {Закрываем запрос}
    If Q_REGION <> nil then begin
       try     If Q_REGION.Active then Q_REGION.Close;
       finally Q_REGION.Free;
       end;
    end;

    {Инициализация}
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
    If Not IS_ADMIN then Exit;

    {Запоминаем пароль}
    WriteGlobalString(INI_SET, INI_SET_PAS_IMPORT_EXPORT, CBPas.Text);
end;


{==============================================================================}
{===================   ACTION: ВОССТАНОВИТЬ НАСТРОЙКИ   =======================}
{==============================================================================}
procedure TFSET.ElPopupButton1Click(Sender: TObject);
begin
    If MessageDlg('Подтвердите сброс всех настроек!'+CH_NEW+'Потребуется перезапуск программы.', mtInformation, [mbYes, mbNo], 0)=mrYes then begin
       Hide;
       Application.ProcessMessages;
       DeleteFile(PATH_WORK_INI);
       Application.Terminate;
    end;
end;


{==============================================================================}
{===================   ACTION: ИЗМЕНИТЬ ГЛАВНЫЙ РЕГИОН   ======================}
{==============================================================================}
procedure TFSET.EMainRegChange(Sender: TObject);
var S: String;
begin
    {Инициализация}
    If (Not (IS_ADMIN or (IMAIN_REGION = 0))) or (EMainReg.Text=SMAIN_REGION) or (EMainReg.Text='') then Exit;

    {Подтверждение}
    If MessageDlg('Подтвердите изменение главного региона на ['+EMainReg.Text+']!'+CH_NEW+'Потребуется перезапуск программы.', mtInformation, [mbYes, mbNo], 0)=mrYes then begin
       {Запоминаем новый регион}
       If Q_REGION.Locate(F_CAPTION, EMainReg.Text, []) then begin
          S := Q_REGION.FieldByName(F_COUNTER).AsString;
       end else begin
          ErrMsg('Недопустимый регион!'); Exit;
       end;
       // Когда несколько регионов с одним именем - ошибка
       // S := ReadFieldFromFilter(@FFMAIN.LBank[0].BD, TABLE_REGIONS, '['+F_CAPTION+']='+QuotedStr(EMainReg.Text), F_COUNTER, true);
       // If Not IsIntegerStr(S) then begin ErrMsg('Недопустимый регион!'); Exit; end;
       WriteGlobalInteger(INI_SET, INI_SET_IMAIN_REGION, StrToInt(S));

       {Завершение программы}
       Hide;
       Application.ProcessMessages;
       Application.Terminate;
    end else begin
       EMainReg.ItemIndex := EMainReg.Items.IndexOf(SMAIN_REGION);
    end;
end;

{==============================================================================}
{===============   ACTION: ИЗМЕНИТЬ ПРИЗНАК РАБОТЫ В СЕТИ   ===================}
{==============================================================================}
procedure TFSET.CBNetKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    CBNetClick(Sender);
end;

procedure TFSET.CBNetClick(Sender: TObject);
begin
    {Инициализация}
    If (Not IS_ADMIN) or (CBNet.Checked=IS_NET) then Exit;

    {Подтверждение}
    If MessageDlg('Подтвердите изменение признака работы в сети!'+CH_NEW+'Потребуется перезапуск программы.', mtInformation, [mbYes, mbNo], 0)=mrYes then begin
       {Запоминаем новый признак}
       WriteGlobalBool(INI_SET, INI_SET_NET, CBNet.Checked);

       {При работе в сети не допускается редактирование кодов т.к. редактируется копия}
       If CBNet.Checked then begin
          CBModifyKod.Checked := false;
          CBModifyKodClick(Sender);
       end;

       {Завершение программы}
       Hide;
       Application.ProcessMessages;
       Application.Terminate;
    end else begin
       CBNet.Checked:=IS_NET;
    end;
end;


{==============================================================================}
{====================   ACTION: ОТКЛЮЧИТЬ ПОЛЬЗОВАТЕЛЕЙ   =====================}
{==============================================================================}
procedure TFSET.CBTerminateClick(Sender: TObject);
begin
    {Инициализация}
    If (Not IS_ADMIN) or (CBTerminate.Checked=IS_TERMINATE) then Exit;

    {Подтверждение}
    If CBTerminate.Checked then begin
       IS_TERMINATE := (MessageDlg('Подтвердите отключение пользователей!', mtInformation, [mbYes, mbNo], 0)=mrYes);
    end else begin
       IS_TERMINATE := false;
    end;

    {Сохраняем настройку}
    CBTerminate.Checked := IS_TERMINATE;
    WriteGlobalBool(INI_SET, INI_SET_TERMINATE, IS_TERMINATE);
end;

procedure TFSET.CBTerminateKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    CBTerminateClick(Sender);
end;


{==============================================================================}
{======================   ЗАКЛАДКА: ОСНОВНЫЕ НАСТРОЙКИ   ======================}
{==============================================================================}
procedure TFSET.CBSEPDetailClick(Sender: TObject);
begin
    {Предупреждение}
    If CBSEPDetail.Checked then begin
       If MessageDlg('Включение оптимизации выбора социально-экономических показателей'+CH_NEW+
                     'значительно увеличивает объем обрабатываемых данных, что может'+CH_NEW+
                     'повлечь замедление расчета, особенно при работе в сети!'+CH_NEW+
                     'Подтвердите операцию!', mtWarning, [mbYes, mbNo], 0) <> mrYes then begin
          CBSEPDetail.Checked := false;
          Exit;
       end;
    end;
    WriteLocalBool(INI_SET, INI_SET_SEP_DETAIL, CBSEPDetail.Checked);
    EnablSEPMonth;
end;

procedure TFSET.CBSEPDetailKeyPress(Sender: TObject; var Key: Char);
begin
    CBSEPDetailClick(Sender);
end;

procedure TFSET.EnablSEPMonth;
begin
    LSepMonth.Enabled  := CBSEPDetail.Checked;
    LSepMonth2.Enabled := CBSEPDetail.Checked;
    TBSEPMonth.Enabled := CBSEPDetail.Checked;
end;

procedure TFSET.TBSEPMonthChange(Sender: TObject);
begin
    LSepMonth2.Caption := IntToStr(TBSEPMonth.Position);
    WriteLocalInteger(INI_SET, INI_SET_SEP_MONTH, TBSEPMonth.Position);
end;



{==============================================================================}
{==========================   ЗАКЛАДКА: ВВОД/ВЫВОД   ==========================}
{==============================================================================}
procedure TFSET.CBImportSubRegClick(Sender: TObject);
begin
    WriteLocalBool(INI_SET, INI_SET_IMPORT_SUBREG, CBImportSubReg.Checked);
end;

procedure TFSET.CBImportSubRegKeyPress(Sender: TObject; var Key: Char);
begin
    CBImportSubRegClick(Sender);
end;

{==============================================================================}

procedure TFSET.CBVerifyRARClick(Sender: TObject);
begin
    WriteLocalBool(INI_SET, INI_SET_IMPORT_VERIFY_RAR, CBVerifyRAR.Checked);
end;

procedure TFSET.CBVerifyRARKeyPress(Sender: TObject; var Key: Char);
begin
    CBVerifyRARClick(Sender);
end;

{==============================================================================}

procedure TFSET.CBVerifyXLSClick(Sender: TObject);
begin
    WriteLocalBool(INI_SET, INI_SET_IMPORT_VERIFY_XLS, CBVerifyXLS.Checked);
end;

procedure TFSET.CBVerifyXLSKeyPress(Sender: TObject; var Key: Char);
begin
    CBVerifyXLSClick(Sender);
end;

{==============================================================================}

procedure TFSET.EPathArjChange(Sender: TObject);
begin
    WriteLocalString(INI_SET, INI_SET_ARJ_PATH, EPathArj.Text);
end;

procedure TFSET.EPathArjRightButtonClick(Sender: TObject);
var SPath : String;
begin
    SPath := EPathArj.Text;
    If SelectDirectory('Каталог архивации данных', '', SPath) then begin
       If SPath[Length(SPath)] <> '\' then SPath:=SPath+'\';
       EPathArj.Text := SPath;
    end;
end;

procedure TFSET.EPathExportChange(Sender: TObject);
begin
    WriteLocalString(INI_SET, INI_SET_EXPORT_PATH, EPathExport.Text);
end;

procedure TFSET.EPathExportRightButtonClick(Sender: TObject);
var SPath : String;
begin
    SPath := EPathExport.Text;
    If SelectDirectory('Каталог экспорта данных', '', SPath) then begin
       If SPath[Length(SPath)] <> '\' then SPath:=SPath+'\';
       EPathExport.Text := SPath;
    end;
end;

{==============================================================================}
{============================   ЗАКЛАДКА: РАЗНОЕ   ============================}
{==============================================================================}
procedure TFSET.CBModifyKodClick(Sender: TObject);
begin
    WriteLocalBool(INI_SET, INI_SET_KOD_MODIFY, CBModifyKod.Checked);
end;

procedure TFSET.CBModifyKodKeyPress(Sender: TObject; var Key: Char);
begin
    CBModifyKodClick(Sender);
end;


end.
