unit MAIN;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.CommCtrl,
  System.SysUtils, System.Variants, System.Classes, System.IniFiles,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ActnList, Vcl.Menus,
  Vcl.StdCtrls, Vcl.XPMan, Vcl.ExtCtrls, Vcl.OleServer, Vcl.ImgList,
  Vcl.ComCtrls, Vcl.ToolWin, Vcl.ActnMan, Vcl.ActnColorMaps,
  Data.DB, Data.Win.ADODB,
  Excel2000, ExcelXP, IdGlobal, IdGlobalProtocols,
  FunType, FunSys, FunInfo;

  // CommCtrl - TVIS_CUT
type
   {Банки данных}
   TBank = record
      ID      : Byte;             // Индекс банка данных: 0, 1, 2, ...
      Pref  : String;             // Префикс банка данных: '', '1', '2', ...
      Caption : String;           // Наименование банка данных
      BD      : TADOConnection;   // База данных индексов
   end;
   TLBank = array of TBank;
   PLBank = ^TLBank;

   {Копия таблицы регионов}
   TRegion = record
      FCounter : Integer;
      FCaption : String;
      FParent  : Integer;
      FGroup   : Integer;                   
      FBegin   : TDate;
      FEnd     : TDate;
   end;
   TLRegion = array of TRegion;
   PLRegion = ^TLRegion;

   {Копия таблицы таблиц}
   TTable = record
      SBankPref : String;        // Префикс банка данных
      FCounter  : String;
      FCaption  : String;
      FNoStand  : Boolean;
      // FCEP   : Boolean;  - не используется
      FBegin    : TDate;
      FEnd      : TDate;
   end;
   TLTable = array of TTable;
   PLTable = ^TLTable;

   {Копия таблицы строк}
   TRow = record
      FNumeric : String;
      FCaption : String;
      FTable   : Integer;
   end;
   TLRow = array of TRow;
   PLRow = ^TLRow;

   {Копия таблицы столбцов}
   TCol = record
      FNumeric : String;
      FCaption : String;
      FTable   : Integer;
   end;
   TLCol = array of TCol;
   PLCol = ^TLCol;

   {Копия таблицы форм}
   TForma = record
      SBankPref   : String;        // Префикс банка данных
      FCounter    : Integer;
      FParent     : Integer;
      FIcon       : Integer;
      FExport     : Boolean;
      FRegion     : String;
      FCaption    : String;
      FFile       : String;
      FMonth1     : Boolean;
      FMonth3     : Boolean;
      FBegin      : TDate;
      FEnd        : TDate;
      FCellID     : String;
      FStrID      : String;
      FCellPeriod : String;
      FCellRegion : String;
   end;
   TLForma = array of TForma;
   PLForma = ^TLForma;

   {Копия таблицы альтернативных кодов}
   TORKey = record
      FKey1 : String;
      FKey2 : String;
   end;
   TLORKey = array of TORKey;
   PLORKey = ^TLORKey;

  TFMAIN = class(TForm)
    ImgLst32: TImageList;
    AList: TActionList;
    AMatrixSet: TAction;
    AHelp: TAction;
    AImport: TAction;
    ImgLst16: TImageList;
    ARefresh: TAction;
    AOpen: TAction;
    ANew: TAction;
    OpenDlg: TOpenDialog;
    XLS: TExcelApplication;
    WB: TExcelWorkbook;
    AOpenTemp: TAction;
    AExport: TAction;
    AMatrixIn: TAction;
    AMatrixOut: TAction;
    ASet: TAction;
    AAbout: TAction;
    AOpenBase: TAction;
    ABlank: TAction;
    ABDRegion: TAction;
    AVerify: TAction;
    ImgLstSys: TImageList;
    ImgLst32BW: TImageList;
    ADel: TAction;
    ADelDir: TAction;
    AFormXLS: TAction;
    AFormINI: TAction;
    TimerShowTime: TTimer;
    ANavPrev12: TAction;
    ANavPrev1: TAction;
    ANavNext1: TAction;
    ANavNext12: TAction;
    ANavPrev3: TAction;
    ANavNext3: TAction;
    AAnalizCompare: TAction;
    AAnalizRating: TAction;
    AClose: TAction;
    TimerInfo: TTimer;
    AInfo: TAction;
    StatusBar: TStatusBar;
    XPManifest1: TXPManifest;
    AFindPrev: TAction;
    AFindNext: TAction;
    Icon: TTrayIcon;
    PMenuServ: TPopupMenu;
    N5: TMenuItem;
    N6: TMenuItem;
    N8: TMenuItem;
    N10: TMenuItem;
    MnuServSeparat1: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    PMenuPath: TPopupMenu;
    N17: TMenuItem;
    N18: TMenuItem;
    PMenuForm: TPopupMenu;
    N27: TMenuItem;
    AKodCorrect: TAction;
    NSet: TMenuItem;
    N7: TMenuItem;
    N9: TMenuItem;
    N19: TMenuItem;
    N20: TMenuItem;
    N21: TMenuItem;
    N22: TMenuItem;
    N23: TMenuItem;
    N24: TMenuItem;
    N25: TMenuItem;
    N26: TMenuItem;
    N2: TMenuItem;
    N1: TMenuItem;
    AKod: TAction;
    N3: TMenuItem;
    AAnalizRegion: TAction;
    PMainRight: TPanel;
    SplitterRight: TSplitter;
    PMain: TPanel;
    CBar: TCoolBar;
    TBar: TToolBar;
    BtnBlank: TToolButton;
    BtnNew: TToolButton;
    BtnImport: TToolButton;
    BtnExport: TToolButton;
    BtnVerify: TToolButton;
    Separator1: TToolButton;
    BtnOpen: TToolButton;
    Separator2: TToolButton;
    BtnServ: TToolButton;
    BtnHelp: TToolButton;
    PMainLeft: TPanel;
    PPathNav: TPanel;
    BtnPrev12: TButton;
    BtnPrev3: TButton;
    BtnPrev1: TButton;
    BtnNext3: TButton;
    BtnNext1: TButton;
    BtnNext12: TButton;
    TreePath: TTreeView;
    SplitterLeft: TSplitter;
    PMainCenter: TPanel;
    PFind: TPanel;
    LFind: TLabel;
    BtnFindPrev: TButton;
    BtnFindNext: TButton;
    CBFind: TComboBox;
    TreeForm: TTreeView;
    PBar: TProgressBar;
    CBarAnaliz: TCoolBar;
    TBarAnaliz: TToolBar;
    ToolButton8: TToolButton;
    ToolButton9: TToolButton;
    ToolButton10: TToolButton;
    PRight: TPanel;
    AAnalizDynamic: TAction;
    ToolButton1: TToolButton;

    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure AImportExecute(Sender: TObject);
    procedure ARefreshExecute(Sender: TObject);
    procedure TreePathChange(Sender: TObject; Node: TTreeNode);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TreeFormDblClick(Sender: TObject);
    procedure AOpenExecute(Sender: TObject);
    procedure TreeFormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ANewExecute(Sender: TObject);
    procedure AOpenTempExecute(Sender: TObject);
    procedure TreePathCompare(Sender: TObject; Node1, Node2: TTreeNode; Data: Integer; var Compare: Integer);
    procedure TreeFormExpanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean);
    procedure TreeFormCollapsing(Sender: TObject; Node: TTreeNode; var AllowCollapse: Boolean);
    procedure AMatrixSetExecute(Sender: TObject);
    procedure AMatrixExecute(Sender: TObject);
    procedure AHelpExecute(Sender: TObject);
    procedure AExportExecute(Sender: TObject);
    procedure ASetExecute(Sender: TObject);
    procedure AAboutExecute(Sender: TObject);
    procedure AOpenBaseExecute(Sender: TObject);
    procedure ABlankExecute(Sender: TObject);
    procedure TreeFormChange(Sender: TObject; Node: TTreeNode);
    procedure ABDRegionExecute(Sender: TObject);
    procedure AVerifyExecute(Sender: TObject);
    procedure ADelExecute(Sender: TObject);
    procedure ADelDirExecute(Sender: TObject);
    procedure TreePathDblClick(Sender: TObject);
    procedure AFormXLSExecute(Sender: TObject);
    procedure AFormINIExecute(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure TimerShowTimeTimer(Sender: TObject);
    procedure TreeFormCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure TreePathCustomDrawItem(Sender: TCustomTreeView; Node: TTreeNode; State: TCustomDrawState; var DefaultDraw: Boolean);
    procedure PPathNavResize(Sender: TObject);
    procedure ANavPathExecute(Sender: TObject);
    procedure AAnalizExecute(Sender: TObject);
    procedure ACloseExecute(Sender: TObject);
    procedure TimerInfoTimer(Sender: TObject);
    procedure AInfoExecute(Sender: TObject);
    procedure WBBeforeClose(ASender: TObject; var Cancel: WordBool);
    procedure WBBeforeSave(ASender: TObject; SaveAsUI: WordBool; var Cancel: WordBool);
    procedure AFindNextExecute(Sender: TObject);
    procedure AFindPrevExecute(Sender: TObject);
    procedure CBFindKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure CBFindChange(Sender: TObject);
    procedure AKodCorrectExecute(Sender: TObject);
    procedure AKodExecute(Sender: TObject);
    procedure TreeFormMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

  private
    IsBlock   : Boolean;       // Признак блокировки ввода
    IsModify  : Boolean;       // Признак модификации книги EXCEL
    IsControl : Boolean;       // Признак контроля открытого отчета
    FindText  : String;        // Искомая строка

    procedure RefreshTreeForm;
    procedure SetBlock;
    procedure ClearBlock;

    procedure ClearSelPath;
    procedure ClearSelForm;

    {МОДУЛЬ:  MVIRTUAL_INI}
    function  IniListBank: Boolean;
    procedure ClearListBank;
    function  IniIndexFile: Boolean;
    function  IniHelpFile: Boolean;
    function  IniListRegions: Boolean;
    function  IniListTables(const IMin, IMax: Integer): Boolean;
    function  IniListForms: Boolean;
    function  IniORKey: Boolean;

  public
    WIN_VER   : TWinVer;        // Версия Windows
    FFPARAM   : TForm;

    LBank     : TLBank;         // Список банков данных
    LRegion   : TLRegion;
    LTable    : TLTable;
    LRow      : TLRow;
    LCol      : TLCol;
    LForma    : TLForma;
    LORKey    : TLORKey;
    LCEP      : TStringList;    // Список таблиц-социально-экономических показателей
    SelPath   : TSelPath;       // Текущее выделение Path
    SelForm   : TSelForm;       // Текущее выделение Form
    IMainBank : Integer;        // Временный приоритетный банк

    procedure EnablAction;
    procedure RefreshTreePath;
    function  SetSelPath: Boolean;
    function  SetSelForm: Boolean;
    function  KeySelForm: String;

    {МОДУЛЬ:  MVIRTUAL}
    function  FormulaToCaption(const SFormula: String; const IsShort: Boolean): String;
    function  KeyToCaption(const SKey, SSeparatKey, SSeparatResult: String; const IsShort: Boolean): String;
    function  RegionsCounterToCaption(const FCounter: Integer): String;
    function  RegionsCounterToParent(const FCounter: Integer): Integer;
    function  TablesCounterToBankPrefix(const FCounter: String; const Dat: TDate): String;
    function  TablesCounterToCaption(const FCounter: String): String;
    function  IsTableNoStandart(const FCounter: String): Boolean;
    function  RowsNumericToCaption(const FNumeric: String; const FTable: Integer): String;
    function  ColsNumericToCaption(const FNumeric: String; const FTable: Integer): String;
    function  FormsCounterToBankPrefix(const FCounter: Integer): String;
    function  FormsCounterToCaption(const FCounter: Integer): String;
    function  FormsCounterToFile(const FCounter: Integer): String;
    function  FormsCounterToParent(const FCounter: Integer): Integer;
    function  FormsCaptionToCounter(const SCaption, SYear, SMonth: String; const OnlyMainBank: Boolean): Integer;
    function  FormsCounterToPathTreeCaption(const FCounter: Integer): String;
    function  IsFormPeriod(const SCaption, SYear, SMonth: String; const Is1, Is3: Boolean): Boolean;

    {МОДУЛЬ:  MAIN_DAT}
    function  PathDat(const SYear, SMonth, SRegion, SKey: String): String;
    function  PathFileForm(const SPref, SFile: String): String;
    function  PathRegion(const SYear, SMonth, SRegion: String): String;
    function  PathToBankPrefix(const SPath: String): String;
    function  BankPrefixToIndex(const SPref: String): Integer;
    function  BankIndexToPrefix(const Ind: Integer): String;

    {МОДУЛЬ:  MAIN_DAT_YMR}
    function  FindDataYMR(const PSr: PArrayStr; const SFindYMR: String; const IsFullPath: Boolean): Integer;
    function  IsTableDatIncludeIniYMR(const SPathIni, YMR: String): Boolean;
    function  SetYMR(const SYear, SMonth, SRegion, SKey: String): String;
    function  IsExistsRowYMR(const YMR, STab, SRow: String): Boolean;
    function  TabToDatYMR(const STab, YMR: String): String;
    function  PrefToDatYMR(const SPref, YMR: String): String;
    function  GetTabKeyYMR(const PList: PStringList; const YMR, Section: String): Boolean;
    function  GetListTabDatYMR(const PList: PStringList; const YMR: String): Boolean;
    function  YMRToDat(const YMR: String): TDate;

    {МОДУЛЬ:  MAIN_DAT_IO}
    function  GetValNoCash(const SPathDat, STab, SCol, SRow: String): String;
    function  IsExistsTab(const SPath, STab: String): Boolean;
    function  IsExistsRow(const SPath, STab, SRow: String): Boolean;
    function  IsTableDatIncludeIni(const SPathIni, SPathDat: String): Boolean;
    function  IsFormReadOnly(const IDForm: Integer): Boolean;

    {МОДУЛЬ:  MAIN_DAT_SECTION}
    function  GetListTabIni(const PList: PStringList; const SPathIni: String): Boolean;
    function  GetListTabDat(const PList: PStringList; const SPathDat: String): Boolean;
    function  GetTabKey(const PList: PStringList; const SPathDat, Section: String): Boolean;
    function  GetTabVal(const SPath, STab: String; const PList: PStringList): Boolean;
    function  SetTabVal(const SPath, STab: String; const PList: PStringList; const IsFullReplace: Boolean): Boolean;
    function  CreateTab(const SPath, STab: String): Boolean;
    function  ClearTab(const SPath, STab: String): Boolean;
    function  DelTab(const SPath, STab: String): Boolean;


    function  FindRegions(const P        : PLRegion; const FindFirst : Boolean;
                          const FCounter : Integer;  const FCaption  : String;
                          const FParent  : Integer;  const FGroup    : Integer;
                          const FDate    : TDate): Boolean;
    function  FindTables(const P: PLTable; const FindFirst: Boolean; const FCounter, FCaption: String; const FDate: TDate): Boolean;
    function  FindRows(const P: PLRow; const FindFirst: Boolean;
                       const FNumeric, FCaption: String; const FTable: Integer): Boolean;
    function  FindCols(const P: PLCol; const FindFirst: Boolean; 
                       const FNumeric, FCaption: String; const FTable: Integer): Boolean;
    function  FindForms(const P: PLForma; const FindFirst, OnlyMainBank: Boolean; const FCounter, FParent: Integer;
                        const FRegion, FCaption, FFile, FMonth: String; const IsID: Boolean; const FDate: TDate): Boolean;
    function  FindORKeys(const P: PLORKey; const FindFirst: Boolean; const FKey1, FKey2: String): Boolean;
    function  Finder(const Directly: Boolean): Boolean;

    function  IsModifyRegion(const SYear, SMonth, SRegion: String; const PLTab: PStringList): Boolean;
    function  ListRegionsAnalitic(const PList: PStringList; const SBaseRegion: String;
                                  const IsApparat, IsMainRegion: Boolean;
                                  const SYear, SMonth: String): Boolean;
    function  ListRegionsRangeUp(const PList: PStringList; const SBaseRegion: String;
                                 const SYear, SMonth: String): Boolean;
    function  ListRegionsRangeDown(const PList: PStringList; const SBaseRegion: String;
                                   const SYear, SMonth: String): Boolean;
    function  ListRegionsLevel(const PList: PStringList; const SBaseRegion: String;
                               const SYear, SMonth: String): Boolean;
    function  CopyListRegions(const PList: PStringList; const SYear, SMonth: String): Boolean;
    function  CreateRegions(const SYear, SMonth, SRegion: String): String;
    function  VerifyRegionName(const SRegionName: String): Boolean;
    function  IsSubRegion(const SRegion: String): Boolean;


    function  RunMacroSafe(const SNameMacros: String): Boolean;
    procedure AlertMsg(const STitle, SMsg: String; const MsgType: TBalloonFlags);
  end;

var
  FMAIN: TFMAIN;

implementation

uses FunConst, FunBD, FunExcel, FunTree, FunDay, FunText, FunIO, FunVcl,
     FunFiles, FunSum, FunVerify, FunIni,
     MEXPORT, MIMPORT_RAR, MIMPORT_XLS, MBASE, MSET, MABOUT, MBLANK, MREPORT,
     MBDREGION, MCLOSE, MTITLE, MSPLASH,
     MPARAM, MKOD, MKOD_CORRECT, MMATRIX, MBANK;

{$R *.dfm}
{$INCLUDE MAIN_ACTION}
{$INCLUDE MAIN_DAT}
{$INCLUDE MAIN_VIRTUAL}
{$INCLUDE MAIN_REGIONS}
{$INCLUDE MAIN_TREE_PATH}
{$INCLUDE MAIN_TREE_FORM}


{==============================================================================}
{============================   СОЗДАНИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFMAIN.FormCreate(Sender: TObject);
var F_SPLASH : TFSPLASH;

    procedure Er;
    var T : TCloseAction;
    begin
        try If (PATH_PROG<>'') then FormClose(Sender, T);
        finally
            Application.ProcessMessages;
            Application.Terminate;
            Halt;
        end;
    end;

begin
    {Запущена ли другая программа}
    If GetProgRun then begin ErrMsg('Программа "Аналитика" уже запущена!'); Er; end;

    {Показываем заставку}
    F_SPLASH := TFSPLASH.Create(Self);
    try
       F_SPLASH.Show;

    {Инициализация}
    F_SPLASH.Msg('5%');
    AAnalizCompare.Tag := TAG_ANALIZ_COMPARE;
    AAnalizRating.Tag  := TAG_ANALIZ_RATING;
    AAnalizRegion.Tag  := TAG_ANALIZ_REGION;
    AAnalizDynamic.Tag := TAG_ANALIZ_DYNAMIC;

    ClearSelPath;
    FFPARAM            := nil;
    IsControl          := false;
    FindText           := '';
    LCEP               := nil;

    PBar.Visible       := false;
    PBar.Min           := 0;
    PBar.Max           := 100;

    {***  Версия Windows  *****************************************************}
    WIN_VER            := GetWinVer;

    {***  Каталог программы  **************************************************}
    PATH_PROG          := ExtractFilePath(Application.EXEName);
    FPATH_PROG_INI     := PATH_PROG+PROG_INI;

    {***  Инициализация рабочего каталога на компьютере пользователя **********}
    PATH_WORK          := GetPathMyDoc;
    If PATH_WORK='' then PATH_WORK := PATH_PROG
                    else PATH_WORK := PATH_WORK+'\';
    PATH_WORK          := PATH_WORK+PATH_MYDOC;
    If Not ForceDirectories(PATH_WORK) then begin ErrMsg('Не могу создать рабочий каталог: '+PATH_WORK); Er; end;

    {*** Инициализация временного каталога ************************************}
    PATH_WORK_TEMP := PATH_WORK+PATH_TEMP;
    DelDir(PATH_WORK_TEMP, true);
    If Not ForceDirectories(PATH_WORK_TEMP) then begin ErrMsg('Ошибка создания каталога временных файлов: '+PATH_WORK_TEMP); Er; end;

    {*** Пути базы данных *****************************************************}
    PATH_BD         := PATH_PROG+PATH_BANK0+'\';
    PATH_BD0        := PATH_PROG+PATH_BANK0;
    PATH_BD_DATA    := PATH_BD+PATH_DATA;
    PATH_BD_SHABLON := PATH_BD+PATH_SHABLON;

    {*** Инициализация списка банков данных ***********************************}
    If Not IniListBank then begin ErrMsg('Не инициализирован список банков данных!'); Er; end;
    F_SPLASH.LBank.Caption := 'Найдено банков данных:  ' + IntToStr(Length(LBank));
    F_SPLASH.LBank.Visible := true;
    F_SPLASH.Msg('10%');

    {*** Инициализация INI-файла локальных настроек ***************************}
    PATH_WORK_INI  := PATH_WORK+WORK_INI;

    {*** Отключение пользователей *********************************************}
    IS_TERMINATE := ReadGlobalBool(INI_SET, INI_SET_TERMINATE, false);

    {*** Режим администратора / пользователя **********************************}
    If DirectoryExists(PATH_BD_DATA) then IS_ADMIN := GetPathWritable(PATH_BD_DATA)
                                     else IS_ADMIN := true;

    //IS_ADMIN := false;  // ОТЛАДКА

    {*** Работа в сети ********************************************************}
    IS_NET := ReadGlobalBool(INI_SET, INI_SET_NET, true);
    If Not IS_ADMIN then IS_NET := true;                  // Для режима пользователя всегда работа в сети
    If IS_NET then StatusBar.Panels[STATUS_NET].Text := 'сеть';

    {*** Инициализация индексных файлов ***************************************}
    F_SPLASH.Msg('15%');
    If Not IniIndexFile then Er;

    {*** Инициализация файла помощи *******************************************}
    F_SPLASH.Msg('25%');
    IniHelpFile;

    {*** Инициализация списка регионов ****************************************}
    F_SPLASH.Msg('35%');
    If Not IniListRegions then Er;

    {*** Инициализация списка таблиц, строк и столбцов ************************}
    F_SPLASH.Msg('50%');
    If Not IniListTables(50, 80) then Er;

    {*** Инициализация списка форм ********************************************}
    F_SPLASH.Msg('80%');
    If Not IniListForms then Er;

    {*** Инициализация альтернативных кодов ***********************************}
    F_SPLASH.Msg('90%');
    If Not IniORKey then Er;

    {Настройка Tree}
    F_SPLASH.Msg('100%');
    TreePath.ControlStyle := TreePath.ControlStyle + [csOpaque];   // Отключение перерисовки фона
    TreeForm.ControlStyle := TreeForm.ControlStyle + [csOpaque];
    TreePath.ToolTips   := false;
    TreeForm.ToolTips   := false;
    TreePath.AutoExpand := true;
    TreeForm.AutoExpand := true;
    TreePath.OnCompare  := TreePathCompare;
    TreePath.PopupMenu  := PMenuPath;
    TreeForm.PopupMenu  := PMenuForm;
    TreePath.Items.Clear;
    TreeForm.Items.Clear;
    TreePath.RightClickSelect := false;
    TreeForm.RightClickSelect := false;

    {Конфигурируем область поиска}
    CBFind.DropDownCount := 25;
    ReadCBListIni(@CBFind, INI_FIND);

    {Настройка EXCEL}
    WB.OnBeforeClose := WBBeforeClose;

    {Настройка Action}
    AMatrixIn.Tag  := 1;
    AMatrixOut.Tag := 2;
    ClearBlock;

    {Конфигурация для администратора}
    If IS_ADMIN then begin
       {Панель инструментов}
       BtnBlank.Visible        := false;
       BtnHelp.Visible         := false;

    {Конфигурация для пользователя}
    end else begin
       {Панель инструментов}
       BtnServ.Visible         := false;
       ANew.Visible            := false;    ANew.Enabled        := false;
       AImport.Visible         := false;    AImport.Enabled     := false;
       AExport.Visible         := false;    AExport.Enabled     := false;
       AVerify.Visible         := false;    AVerify.Enabled     := false;

       {Меню: Сервис}
       ASet.Visible            := false;    ASet.Enabled        := false;
       NSet.Visible            := false;    NSet.Enabled        := false;
       AMatrixIn.Visible       := false;    AMatrixIn.Enabled   := false;
       AMatrixOut.Visible      := false;    AMatrixOut.Enabled  := false;
       AMatrixSet.Visible      := false;    AMatrixSet.Enabled  := false;
       AFormXLS.Visible        := false;    AFormXLS.Enabled    := false;
       AFormINI.Visible        := false;    AFormINI.Enabled    := false;
       AKodCorrect.Visible     := false;    AKodCorrect.Enabled := false;
       ABDRegion.Visible       := false;    ABDRegion.Enabled   := false;
       AOpenBase.Visible       := false;    AOpenBase.Enabled   := false;
       AOpenTemp.Visible       := false;    AOpenTemp.Enabled   := false;
       AInfo.Visible           := false;    AInfo.Enabled       := false;

       {Меню: TreePath}
       ADelDir.Visible         := false;    ADelDir.Enabled     := false;

       {Меню: TreeForm}
       ADel.Visible            := false;    ADel.Enabled        := false;

       Repaint;
    end;

    {Настройка иконки-информатора}
    Icon.Visible               := false;

       F_SPLASH.Hide;
    finally
       F_SPLASH.Free;
    end;

    {Запускаем таймер информатора}
    TimerInfo.Interval := TIMER_INFO_1;
    TimerInfo.Enabled  := true;
end;


{==============================================================================}
{==============================   ПОКАЗ ФОРМЫ    ==============================}
{==============================================================================}
procedure TFMAIN.FormShow(Sender: TObject);
begin
    {Восстанавливаем настройки из Ini}
    LoadFormIni(Self, [fspPosition, fspState]);
    PMainLeft.Width  := ReadLocalInteger(INI_FORM_PARAM+Self.Name, INI_FORM_PARAM_SEP_LEFT,  300);
    PMainRight.Width := ReadLocalInteger(INI_FORM_PARAM+Self.Name, INI_FORM_PARAM_SEP_RIGHT, 500);

    {Устраняем глюк с высотой панели}
    TBar.Visible := false;
    TBar.Visible := true;

    {Востанавливаем структуры деревьев}
    RefreshTreePath;

    {Загрузка аналитики}
    FFPARAM := TFPARAM(LoadSubForm(PRight, TFPARAM, true));

    {Окно на передний план}
    ForegroundWindow;
end;


{==============================================================================}
{==========================   ЗАПРОС НА ВЫХОД    ==============================}
{==============================================================================}
procedure TFMAIN.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var FClose: TFCLOSE;
begin
    {Подтверждение на выход}
    If Sender<>nil then begin
       FCLOSE := TFCLOSE.Create(Self);
       try     CanClose := FCLOSE.Execute;
               If Not CanClose then Exit;
       finally FCLOSE.Free;
       end;
    end;

    {Сохраняем настройки в Ini}
    SaveFormIni(Self, [fspPosition, fspState]);
    WriteLocalInteger(INI_FORM_PARAM+Self.Name, INI_FORM_PARAM_SEP_LEFT,  PMainLeft.Width);
    WriteLocalInteger(INI_FORM_PARAM+Self.Name, INI_FORM_PARAM_SEP_RIGHT, PMainRight.Width);
end;


{==============================================================================}
{============================   ЗАКРЫТИЕ ФОРМЫ    =============================}
{==============================================================================}
procedure TFMAIN.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {Останавливаем таймеры}
    TimerShowTime.Enabled := false;
    TimerInfo.Enabled     := false;

    {Выгружаем форму аналитики}
    FFPARAM := LoadSubForm(PRight, nil, true);

    {Освобождаем память}
    ClearListBank;
    SetLength(LRegion, 0);
    SetLength(LTable,  0);
    SetLength(LRow,    0);
    SetLength(LCol,    0);
    SetLength(LForma,  0);
    SetLength(LORKey,  0);
    try If LCEP<>nil then LCEP.Free; except end;

    {Удаляем временные файлы в рабочем каталоге}
    DelDir(PATH_WORK_TEMP, true);
end;


{==============================================================================}
{=========================   ACTION: НОВЫЙ ОТЧЕТ    ===========================}
{==============================================================================}
procedure TFMAIN.ANewExecute(Sender: TObject);
var F_REPORT : TREPORT;
    F_TITLE  : TFTITLE;
    BPrev    : Boolean;
    IPrev    : Integer;
    SYear, SMonth, SRegion, SForm : String;
begin
    {Инициализация}
    If (Not ANew.Enabled) or (Not IS_ADMIN) then Exit;

    {Период по умолчанию}
    If Not NullPeriod(SYear, SMonth) then Exit;

    {Подверждение-корректировка периода, региона и формы}
    F_TITLE := TFTITLE.Create(nil);
    try     If Not F_TITLE.Execute(SYear, SMonth, SRegion, SForm, true, true) then Exit;
            BPrev := F_TITLE.CBPrev.Checked;
    finally F_TITLE.Free;
    end;

    {Создание отчета}
    IsControl:=false;
    SetBlock;
    IPrev := 0;
    try
       F_REPORT := TREPORT.Create;
       try
          {Новый отчет с данными за предыдущий период}
          If BPrev then begin
             If IsFormPeriod(SForm, SYear, SMonth, true,  false) then IPrev := -1;
             If IsFormPeriod(SForm, SYear, SMonth, false, true)  then IPrev := -3;
          end;
          IsControl := F_REPORT.CreateReport(SYear, SMonth, SRegion, '', FormsCaptionToCounter(SForm, SYear, SMonth, true), 0, IPrev, true);
       finally
          F_REPORT.Free;
       end;
    finally
       {Режим ожидания если отчет создан}
       If IsControl then begin
          IsModify := true;
          Enabled  := false;
       end else begin
          ClearBlock;
       end;
    end;
end;


{==============================================================================}
{========================   ACTION: ОТКРЫТЬ ОТЧЕТ    ==========================}
{==============================================================================}
procedure TFMAIN.AOpenExecute(Sender: TObject);
var F_REPORT : TREPORT;
    NForm    : TTreeNode;
    SPath    : String;
    FBANK    : TFBANK;
begin
    {Инициализация}
    If Not AOpen.Enabled then Exit;
    If Not SetSelPath then Exit;
    NForm:=TreeForm.Selected;
    If NForm=nil then Exit;
    If Integer(NForm.Data) < 0 then Exit;

    {Произвольный файл}
    {!!! Важно !!! - кривая работа при отчете с уровнем 0}
    If (NForm.Level=0) and (FindIntInArray(NForm.ImageIndex, [ICO_FORM_FREE_XLS, ICO_FORM_FREE_PDF, ICO_FORM_FREE_MDI])) then begin
       {Определяем файл}
       SPath := PATH_BD0+BankIndexToPrefix(Integer(NForm.Data))+'\'+PATH_DATA+SelPath.SPath+NForm.Text;
       Case NForm.ImageIndex of
       ICO_FORM_FREE_XLS: SPath := SPath + '.xls';
       ICO_FORM_FREE_PDF: SPath := SPath + '.pdf';
       ICO_FORM_FREE_MDI: SPath := SPath + '.mdi';
       else Exit;
       end;
       {Пытаемся устранить кривую работу: если это отчет, то файла не будет}
       If FileExists(SPath) then begin
          If Not IS_ADMIN then SPath := CopyFileToTemp(SPath);
          If SPath='' then begin ErrMsg('Не определен файл отчета!'+CH_NEW+SPath); Exit; end;
          {Открываем файл}
          StartAssociatedExe(SPath, SW_SHOWMAXIMIZED);
          Exit;
       end;
    end;

    {Заполнение отчета}
    IsControl:=false;
    SetBlock;
    try
       F_REPORT := TREPORT.Create;
       try
          Case NForm.ImageIndex of
          ICO_FORM_DETAIL_COL: IsControl := F_REPORT.CreateReport(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, KeySelForm, IDFORM_VIEW_COL, 0, 0, false); {Открыть значение}
          else begin
                  {Выбор индекса главного банка данных}
                  If GetKeyState(VK_LCONTROL) then begin
                     FBANK := TFBANK.Create(Self);
                     try     IMainBank := FBANK.Execute;
                     finally FBANK.Free;
                     end;
                     If IMainBank < -1 then Exit;
                  end;
                  IsControl := F_REPORT.CreateReport(SelPath.SYear, SelPath.SMonth, SelPath.SRegion, '', Integer(NForm.Data), 0, 0, false);
               end;
          end;
       finally
          F_REPORT.Free;
          {Сбрасываваем индекс главного банка данных}
          IMainBank := -1;
       end;
    finally
       {Режим ожидания если отчет создан}
       If IsControl and (XLS.Workbooks.Count > 0) then begin
          IsModify := false;
          Enabled  := false;
       end else begin
          ClearBlock;
       end;
    end;
end;


{==============================================================================}
{======================   КОНТРОЛЬ СОХРАНЕНИЯ ОТЧЕТА    =======================}
{==============================================================================}
procedure TFMAIN.WBBeforeSave(ASender: TObject; SaveAsUI: WordBool; var Cancel: WordBool);
begin
    IsModify:=true;
end;


{==============================================================================}
{=======================   КОНТРОЛЬ ЗАКРЫТИЯ ОТЧЕТА    ========================}
{==============================================================================}
procedure TFMAIN.WBBeforeClose(ASender: TObject; var Cancel: WordBool);
var FIMPORT_XLS               : TIMPORT_XLS;
    LTab                      : TStringList;
    S, SYear, SMonth, SRegion : String;
begin
    {Инициализация}
    Enabled := true;
    try
       {Сохраняем позицию курсора}
       SaveActiveCell(XLS.Workbooks[1]);

       {Делаем отчет невидимым - ВОЗМОЖНЫ СБОИ}
       // try WB.Windows.Item[1].Visible := false; except end;

       {Импортируем отчет}
       If  IsControl and (IsModify Or (Not XLS.Workbooks[1].Saved[LOCALE_USER_DEFAULT])) then begin                            // возможность записи
          {Отключаем признак контроля}
          IsControl:=false;

          {Окно программы на передний план}
          ForegroundWindow;

          {Имеет регион структурные подразделения}
          SYear   := Variant(XLS.Workbooks[1]).WorkSheets[1].Range[CELL_YEAR].Value;
          SMonth  := Variant(XLS.Workbooks[1]).WorkSheets[1].Range[CELL_MONTH].Value;
          SRegion := Variant(XLS.Workbooks[1]).WorkSheets[1].Range[CELL_REGION].Value;

          {Список контролируемых таблиц}
          LTab:=TStringList.Create;
          try
             S := Variant(XLS.Workbooks[1]).WorkSheets[1].Range[CELL_LOG].Value;
             Delete(S, 1, Length(IDLOG));
             S := Trim(S);
             S := FormsCounterToFile(StrToInt(S));
             S := ChangeFileExt(S, '.ini');
             If Not GetListTabIni(@LTab, S) then LTab.Clear;

             {Запись отчета в банк данных}
             If IsModifyRegion(SYear, SMonth, SRegion, @LTab) then begin
                If MessageDlg('Сохранить информацию в базе данных?', mtConfirmation, [mbYes, mbNo], 0)=mrYes then begin
                   FIMPORT_XLS:=TIMPORT_XLS.Create;
                   try     If Not FIMPORT_XLS.ImportXlsID(XLS.Workbooks[1], false) then ErrMsg('Ошибка импорта данных!');
                   finally FIMPORT_XLS.Free;
                   end;
                end;

             {Блокировка записи отчета в банк данных}
             end else begin
                MessageDlg('В целях обеспечения целостности данных сохранение заблокировано.'+CH_NEW+
                           'Изменения необходимо вносить в отчеты структурных подразделений!', mtInformation, [mbOk], 0);
             end;
          finally
             LTab.Free;
          end;

          {Обновляем списки}
          Refresh;
          RefreshTreePath;
       end;

       {Закрытие EXCEL}
       Refresh;
       XLS.DisplayAlerts[0]:=false;
       //XLS.DisplayAlerts[1]:=false; // ошибка в Office 2010
       XLS.Workbooks[1].Saved[LOCALE_USER_DEFAULT]:=true;
       WB.Disconnect;
       XLS.UserControl := True;
       XLS.Quit;
       XLS.Disconnect;
    finally
       {Удаляем блокировку}
       ClearBlock;
    end;
end;


{==============================================================================}
{========================   ACTION: ИМПОРТ ОТЧЕТА    ==========================}
{==============================================================================}
procedure TFMAIN.AImportExecute(Sender: TObject);
var FIMPORT_RAR : TFIMPORT_RAR;
    FIMPORT_XLS : TIMPORT_XLS;
    SPath       : String;
    SExt        : String;
    I           : Integer;
begin
    {Допустимость}                         
    If (Not AImport.Enabled) or (Not IS_ADMIN) then Exit;

    {Устанавливаем блокировку}
    SetBlock;
    try
       {Имя отчета}
       If Not OpenDlg.Execute then Exit;
       Repaint;

       {Просматривает список файлов}
       For I:=0 to OpenDlg.Files.Count-1 do begin

          {Выбираем очередной файл}
          SPath := OpenDlg.Files[I];
          If Not FileExists(SPath) then Continue;
          SExt:=ExtractFileExt(SPath);

          {Импорт архива RAR}
          If CmpStr(SExt, '.rar') then begin
             FIMPORT_RAR:=TFIMPORT_RAR.Create(Self);
             try     FIMPORT_RAR.Execute(SPath);
             finally FIMPORT_RAR.Free;
             end;
          end;

          {Импорт отчета EXCEL}
          If CmpStr(SExt, '.xls') then begin
             FIMPORT_XLS:=TIMPORT_XLS.Create;
             try     If Not FIMPORT_XLS.ImportXls(SPath) then Exit;
             finally FIMPORT_XLS.Free;
             end;
          end;

          {Обновляем списки}
          RefreshTreePath;
       end;

    finally
       {Удаляем блокировку}
       ClearBlock;
    end;                    
end;


procedure TFMAIN.SetBlock;
begin
    IsBlock:=true;
    EnablAction;
end;

procedure TFMAIN.ClearBlock;
begin
    IsBlock:=false;
    EnablAction;
end;


{==============================================================================}
{======================   ЗАПУСК МАКРОСА БЕЗ ОШИБКИ    ========================}
{==============================================================================}
function TFMAIN.RunMacroSafe(const SNameMacros: String): Boolean;
var IsAlert : Boolean;
begin
    Result:=false;
    IsAlert := XLS.DisplayAlerts[0];
    try
       XLS.DisplayAlerts[0] := false;
       XLS.Run(SNameMacros);
       Result := true;
    except
       XLS.DisplayAlerts[0] := IsAlert;
    end;
end;


{==============================================================================}
{==============================   СООБЩЕНИЕ    ================================}
{==============================================================================}
procedure TFMAIN.AlertMsg(const STitle, SMsg: String; const MsgType: TBalloonFlags);
begin
    {Инициализация}
    If Trim(SMsg)='' then Exit;
    Icon.Animate      := true;
    Icon.Visible      := true;

    {Добавляем сообщение}
    Icon.BalloonFlags := MsgType;
    Icon.BalloonTitle := STitle;
    Icon.BalloonHint  := SMsg;

    {Показываем сообщение}
    Icon.ShowBalloonHint;
end;



procedure TFMAIN.TimerShowTimeTimer(Sender: TObject);
begin
    TimerShowTime.Enabled := false;
    try
      StatusBar.Panels.BeginUpdate;
      try StatusBar.Panels[STATUS_TIME].Text:=' ' + FormatDateTime('dd mmmm yyyy г.    hh:mm:ss', Now)+'      ';
      except end;
    finally StatusBar.Panels.EndUpdate;
    end;
    TimerShowTime.Enabled := true;
end;

procedure TFMAIN.TimerInfoTimer(Sender: TObject);
begin
    {Инициализация}
    TimerInfo.Enabled  := false;
    try
       {*** Периодически ******************************************************}
       If TimerInfo.Interval = TIMER_INFO_2 then begin
          Icon.Visible := false;
          If IS_ADMIN then Exit;
       end;

       {***  Только при запуске  **********************************************}
       If TimerInfo.Interval = TIMER_INFO_1 then begin
          TimerInfo.Interval := TIMER_INFO_2;

          {Сообщение информатора}
          If Not IS_ADMIN then AlertMsg('Важно!', ReadTxtFile(PATH_BD+FILE_LOCAL_INFO), bfInfo);
       end;

       {*** Всегда ************************************************************}
       IS_TERMINATE := ReadGlobalBool(INI_SET, INI_SET_TERMINATE, false);
       If IS_TERMINATE then begin
          If IS_ADMIN then begin
             AlertMsg('Внимание!', 'Включен режим блокировки пользователей!',  bfWarning);
          end else begin
             AlertMsg('Внимание!', 'По требованию администратора работа программы "АНАЛИТИКА" завершена!', bfError);
             Application.ProcessMessages;
             Sleep(3000);
             Application.ProcessMessages;
             Application.Terminate;
             Exit;
          end;
       end;

       {Запускаем таймер}
       TimerInfo.Enabled := true;
    except
       Icon.Visible := false;
    end;
end;


{==============================================================================}
{===================   СОБЫТИЕ: ИЗМЕНЕНИЕ РАЗМЕРОВ ОКНА    ====================}
{==============================================================================}
procedure TFMAIN.FormResize(Sender: TObject);
begin
    SaveFormIni(Self, [fspPosition, fspState]);
    StatusBar.Update;
    StatusBar.Panels[STATUS_MAIN].Width := StatusBar.ClientWidth - StatusBar.Panels[STATUS_KOD].Width
                                                                 - StatusBar.Panels[STATUS_NET].Width
                                                                 - StatusBar.Panels[STATUS_REGION].Width
                                                                 - StatusBar.Panels[STATUS_TIME].Width;
end;

end.
