unit MKOD;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.DBCtrls, Vcl.Grids,
  Vcl.ActnList, Vcl.Buttons, Vcl.DBGrids, Vcl.StdCtrls, Vcl.ExtCtrls, Vcl.ComCtrls,
  Data.DB, Data.Win.ADODB,
  MAIN;

type
  PDBGrid      = ^TDBGrid;
  PDBNavigator = ^TDBNavigator;

  TKOD = record
     STab, SCol, SRow: String;
  end;

  TFKOD = class(TForm)
    T_TAB: TADOTable;
    T_ROW: TADOTable;
    T_COL: TADOTable;
    DS_TAB: TDataSource;
    DS_ROW: TDataSource;
    DS_COL: TDataSource;
    PControl: TPageControl;
    TS_Tab: TTabSheet;
    TS_Col: TTabSheet;
    TS_Row: TTabSheet;
    BCheck1: TBevel;
    CBModify: TCheckBox;
    Bevel2: TBevel;
    PTab: TPanel;
    LTabCaption: TLabel;
    LTab: TDBText;
    ETab: TDBText;
    PCol: TPanel;
    PRow: TPanel;
    LColCaption: TLabel;
    LRowCaption: TLabel;
    LCol: TDBText;
    LRow: TDBText;
    ECol: TDBText;
    ERow: TDBText;
    Bevel3: TBevel;
    GTab: TDBGrid;
    Bevel4: TBevel;
    Bevel5: TBevel;
    Bevel6: TBevel;
    Bevel7: TBevel;
    Bevel8: TBevel;
    Bevel9: TBevel;
    NTab: TDBNavigator;
    NCol: TDBNavigator;
    NRow: TDBNavigator;
    GCol: TDBGrid;
    GRow: TDBGrid;
    PBottom: TPanel;
    BtnOk: TBitBtn;
    AList: TActionList;
    AOk: TAction;
    AClose: TAction;
    ACancel: TAction;
    BtnClose: TBitBtn;
    BtnCancel: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure T_AfterScroll(DataSet: TDataSet);
    procedure TS_Resize(Sender: TObject);
    procedure CBModifyClick(Sender: TObject);
    procedure CBModifyKeyPress(Sender: TObject; var Key: Char);
    procedure FormActivate(Sender: TObject);
    procedure T_TABAfterScroll(DataSet: TDataSet);
    procedure TabClick(Sender: TObject);
    procedure ColClick(Sender: TObject);
    procedure RowClick(Sender: TObject);
    procedure CloseOk(Sender: TObject);
    procedure CloseCancel(Sender: TObject);
    procedure T_ROWNewRecord(DataSet: TDataSet);
    procedure T_COLNewRecord(DataSet: TDataSet);
  private
    FFMAIN : TFMAIN;
    KOD    : TKOD;

    procedure CorrectInterface;
    procedure SaveColWidth;
    procedure LoadColWidth;
  public
    function  Execute(const SKod: String; const IsOneBtnClose: Boolean): String;
  end;

var
  FKOD: TFKOD;

implementation

uses FunConst, FunSys, FunIni, FunBD, FunText;

{$R *.dfm}

{==============================================================================}
{===========================   СОЗДАНИЕ ФОРМЫ   ===============================}
{==============================================================================}
procedure TFKOD.FormCreate(Sender: TObject);
var Col: TColumn;
begin
    {Инициализация}
    FFMAIN      := TFMAIN(GlFindComponent('FMAIN'));
    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Выбор параметра';

    {Восстанавливаем настройки из Ini}
    LoadFormIni(Self, [fspPosition]);

    {Признак режима модификации}
    If IS_ADMIN and (Not IS_NET) then begin
       CBModify.Checked:=ReadLocalBool(INI_SET, INI_SET_KOD_MODIFY, false);
    end else begin
       CBModify.Visible := false;
       CBModify.Enabled := false;
       CBModify.Checked := false;
       BCheck1.Visible  := false;
    end;

    KOD.STab        := '';
    KOD.SCol        := '';
    KOD.SRow        := '';

    LTab.DataSource := DS_TAB;
    LRow.DataSource := DS_ROW;
    LCol.DataSource := DS_COL;

    LTab.DataField  := F_COUNTER;
    LRow.DataField  := F_NUMERIC;
    LCol.DataField  := F_NUMERIC;

    ETab.DataSource := DS_TAB;
    ERow.DataSource := DS_ROW;
    ECol.DataSource := DS_COL;

    ETab.DataField  := F_CAPTION;
    ERow.DataField  := F_CAPTION;
    ECol.DataField  := F_CAPTION;

    NTab.DataSource := DS_TAB;
    NRow.DataSource := DS_ROW;
    NCol.DataSource := DS_COL;

    GTab.DataSource := DS_TAB;
    GRow.DataSource := DS_ROW;
    GCol.DataSource := DS_COL;

    DS_TAB.DataSet  := T_TAB;
    DS_ROW.DataSet  := T_ROW;
    DS_COL.DataSet  := T_COL;

    T_TAB.AfterScroll := T_TABAfterScroll;
    T_ROW.AfterScroll := T_AfterScroll;
    T_COL.AfterScroll := T_AfterScroll;

    AOk.OnExecute          := CloseOk;
    AClose.OnExecute       := CloseOk;
    ACancel.OnExecute      := CloseCancel;

    PTab.OnDblClick        := CloseOk;
    LTabCaption.OnDblClick := CloseOk;
    LTab.OnDblClick        := CloseOk;
    ETab.OnDblClick        := CloseOk;
    // GTab.OnDblClick     := CloseOk;

    PCol.OnDblClick        := CloseOk;
    LColCaption.OnDblClick := CloseOk;
    LCol.OnDblClick        := CloseOk;
    ECol.OnDblClick        := CloseOk;
    // GCol.OnDblClick     := CloseOk;

    PRow.OnDblClick        := CloseOk;
    LRowCaption.OnDblClick := CloseOk;
    LRow.OnDblClick        := CloseOk;
    ERow.OnDblClick        := CloseOk;
    // GRow.OnDblClick     := CloseOk;

    PTab.OnClick        := TabClick;      PTab.Cursor        := crHandPoint;
    LTabCaption.OnClick := TabClick;      LTabCaption.Cursor := crHandPoint;
    LTab.OnClick        := TabClick;      LTab.Cursor        := crHandPoint;
    ETab.OnClick        := TabClick;      ETab.Cursor        := crHandPoint;

    PCol.OnClick        := ColClick;      PCol.Cursor        := crHandPoint;
    LColCaption.OnClick := ColClick;      LColCaption.Cursor := crHandPoint;
    LCol.OnClick        := ColClick;      LCol.Cursor        := crHandPoint;
    ECol.OnClick        := ColClick;      ECol.Cursor        := crHandPoint;

    PRow.OnClick        := RowClick;      PRow.Cursor        := crHandPoint;
    LRowCaption.OnClick := RowClick;      LRowCaption.Cursor := crHandPoint;
    LRow.OnClick        := RowClick;      LRow.Cursor        := crHandPoint;
    ERow.OnClick        := RowClick;      ERow.Cursor        := crHandPoint;

    {Инициализация таблиц}
    OpenBD(@FFMAIN.LBank[0].BD, '', '', [@T_TAB,       @T_COL,     @T_ROW],
                                        [TABLE_TABLES, TABLE_COLS, TABLE_ROWS]);

    {Сортировка ОБЯЗАТЕЛЬНО до связи таблиц}
    T_TAB.Sort := '['+F_COUNTER+'] ASC';
    T_COL.Sort := '['+F_NUMERIC+'] ASC';
    T_ROW.Sort := '['+F_NUMERIC+'] ASC';

    {Формируем столбцы}
    GTab.Columns.Clear;
    Col             := GTab.Columns.Add;
    Col.FieldName   := T_TAB.Fields[NTABLES_COUNTER].FieldName;
    Col.Width       := 40;
    Col.Font.Color  := ClRed;

    Col             := GTab.Columns.Add;
    Col.FieldName   := T_TAB.Fields[NTABLES_CAPTION].FieldName;
    Col.Font.Color  := ClBlue;

    Col             := GTab.Columns.Add;
    Col.FieldName   := T_TAB.Fields[NTABLES_NOSTAND].FieldName;
    Col.Width       := 40;

    Col             := GTab.Columns.Add;
    Col.FieldName   := T_TAB.Fields[NTABLES_CEP].FieldName;
    Col.Width       := 40;

    Col             := GTab.Columns.Add;
    Col.FieldName   := T_TAB.Fields[NTABLES_BEGIN].FieldName;
    Col.Width       := 70;

    Col             := GTab.Columns.Add;
    Col.FieldName   := T_TAB.Fields[NTABLES_END].FieldName;
    Col.Width       := 70;

    Col            := GCol.Columns.Add;
    Col.FieldName  := F_NUMERIC;
    Col.Width      := 50;
    Col.Font.Color := ClRed;

    Col            := GCol.Columns.Add;
    Col.FieldName  := F_CAPTION;
    Col.Font.Color := ClBlue;

    Col            := GRow.Columns.Add;
    Col.FieldName  := F_NUMERIC;
    Col.Width      := 50;
    Col.Font.Color := ClRed;

    Col            := GRow.Columns.Add;
    Col.FieldName  := F_CAPTION;
    Col.Font.Color := ClBlue;

    {Загружаем ширину столбцов}
    LoadColWidth;

    {Доступность модификации}
    CBModifyClick(Sender);
end;


{==============================================================================}
{========================       ВНЕШНЯЯ ФУНКЦИЯ       =========================}
{==============================================================================}
{========================    SKOD = STAB/SCOL/SROW    =========================}
{==============================================================================}
function TFKOD.Execute(const SKod: String; const IsOneBtnClose: Boolean): String;
var STab, SCol, SRow : String;
begin
    {Инициализация}
    Result := SKod;
    BtnOk.Visible     := Not IsOneBtnClose;
    BtnCancel.Visible := Not IsOneBtnClose;
    BtnClose.Visible  := IsOneBtnClose;

    {Разделяем код}
    If Not SeparatKod(SKod, STab, SRow, SCol) then Exit;

    {Позиционируем таблицы}
    T_ROW.IndexFieldNames := F_NUMERIC;
    T_COL.IndexFieldNames := F_NUMERIC;
    If STab <> '' then T_TAB.Locate(F_COUNTER, STab, []) else T_TAB.First;
    If SRow <> '' then T_ROW.Locate(F_NUMERIC, StrToInt(SRow), []) else T_ROW.First;
    If SCol <> '' then T_COL.Locate(F_NUMERIC, StrToInt(SCol), []) else T_COL.First;

    {Показ формы}
    If ShowModal = mrOk then begin
       Result:=GeneratKod(KOD.STab, KOD.SRow, KOD.SCol);
    end;
end;


{==============================================================================}
{=========================   АКТИВАЦИЯ ФОРМЫ   ================================}
{==============================================================================}
procedure TFKOD.FormActivate(Sender: TObject);
begin
    GTab.SetFocus;
end;


{==============================================================================}
{=======================   АКТИВАЦИЯ ЗАКЛАДОК   ===============================}
{==============================================================================}
procedure TFKOD.TabClick(Sender: TObject);
begin PControl.ActivePageIndex := 0; end;

procedure TFKOD.ColClick(Sender: TObject);
begin PControl.ActivePageIndex := 1; end;

procedure TFKOD.RowClick(Sender: TObject);
begin PControl.ActivePageIndex := 2; end;


{==============================================================================}
{==========================   ЗАКРЫТИЕ ФОРМЫ   ================================}
{==============================================================================}
procedure TFKOD.FormClose(Sender: TObject; var Action: TCloseAction);
begin
    {Сохраняем настройки в Ini}
    SaveFormIni(Self, [fspPosition]);

    {Сохраняем ширину столбцов}
    SaveColWidth;

    {Закрываем таблицы}
    CloseBD(nil, [@T_TAB, @T_COL, @T_ROW]);

    FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:='';
end;


{==============================================================================}
{============================   СОБЫТИЕ mrOk   ================================}
{==============================================================================}
procedure TFKOD.CloseOk(Sender: TObject);
begin ModalResult:=mrOk; end;

{==============================================================================}
{============================ СОБЫТИЕ mrCancel ================================}
{==============================================================================}
procedure TFKOD.CloseCancel(Sender: TObject);
begin ModalResult:=mrCancel; end;


{==============================================================================}
{========================   СОБЫТИЕ OnNewRecord   =============================}
{==============================================================================}
procedure TFKOD.T_COLNewRecord(DataSet: TDataSet);
begin
    T_COL.DisableControls;
    T_COL.FieldByName(TABLE_TABLES).AsInteger := T_TAB.FieldByName(F_COUNTER).AsInteger;
    T_COL.EnableControls;
end;

procedure TFKOD.T_ROWNewRecord(DataSet: TDataSet);
begin
    T_ROW.DisableControls;
    T_ROW.FieldByName(TABLE_TABLES).AsInteger := T_TAB.FieldByName(F_COUNTER).AsInteger;
    T_ROW.EnableControls;
end;

{==============================================================================}
{=============   ПРИ ИЗМЕНЕНИИ ДАННЫХ КОРРЕКТИРОВКА ИНТЕРФЕЙСА   ==============}
{==============================================================================}
procedure TFKOD.T_TABAfterScroll(DataSet: TDataSet);
var SInd: String;
begin
    {Устанавливаем новые выборки строк и столбцов}
    SInd := T_TAB.FieldByName(F_COUNTER).AsString;
    If SInd <> '' then begin
       SetDBFilter(@T_ROW, '['+TABLE_TABLES+']='+SInd);
       SetDBFilter(@T_COL, '['+TABLE_TABLES+']='+SInd);
    end;

    {Корректируем интерфейс}
    CorrectInterface;
end;

procedure TFKOD.T_AfterScroll(DataSet: TDataSet);
begin
    {Корректируем интерфейс}
    CorrectInterface;
end;


{==============================================================================}
{========================   КОРРЕКТИРОВКА ИНТЕРФЕЙСА   ========================}
{==============================================================================}
procedure TFKOD.CorrectInterface;
begin
    {Инициализация}
    If (Not T_TAB.Active) or (Not T_COL.Active) or (Not T_ROW.Active) then begin
       AOk.Enabled:=false;
       Exit;
    end;

    {Доступность кнопки OK}
    AOk.Enabled := (T_TAB.RecordCount > 0) and
                   (T_COL.RecordCount > 0) and
                   (T_ROW.RecordCount > 0);

    {Текущий код}
    If T_TAB.RecordCount>0 then KOD.STab := T_TAB.FieldByName(F_COUNTER).AsString
                           else KOD.STab := '';
    If T_COL.RecordCount>0 then KOD.SCol := T_COL.FieldByName(F_NUMERIC).AsString
                           else KOD.SCol := '';
    If T_ROW.RecordCount>0 then KOD.SRow := T_ROW.FieldByName(F_NUMERIC).AsString
                           else KOD.SRow := '';
    Caption := 'Код параметра: '+GeneratKod(KOD.STab, KOD.SRow, KOD.SCol);
end;

{==============================================================================}
{=======================   ИЗМЕНЕНИЕ РЕЖИМА ЗАПИСИ    =========================}
{==============================================================================}
procedure TFKOD.CBModifyClick(Sender: TObject);
    procedure SetMod(const PGrid: PDBGrid; const PNav: PDBNavigator);
    begin
        If CBModify.Checked and IS_ADMIN then begin
           PNav^.VisibleButtons := [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete];
           PGrid^.Options       := PGrid^.Options + [dgEditing] - [dgRowSelect];
           PGrid^.ReadOnly      := false;
        end else begin
           PNav^.VisibleButtons := [nbFirst, nbPrior, nbNext, nbLast];
           PGrid^.Options       := PGrid^.Options - [dgEditing] + [dgRowSelect];
           PGrid^.ReadOnly      := true;
        end;
    end;
begin
    {Запоминаем установку}
    If IS_ADMIN and (Not IS_NET) then WriteLocalBool(INI_SET, INI_SET_KOD_MODIFY, CBModify.Checked);

    {Изменяем настройки}
    SetMod(@GTab, @NTab);
    SetMod(@GCol, @NCol);
    SetMod(@GRow, @NRow);
end;

procedure TFKOD.CBModifyKeyPress(Sender: TObject; var Key: Char);
begin
    CBModifyClick(Sender);
end;


{==============================================================================}
{========================    СОБЫТИЕ: ON_RESIZE    ============================}
{==============================================================================}
procedure TFKOD.TS_Resize(Sender: TObject);
var  P : PDBGrid;
     S : String;
     I, IWidth, ICaption : Integer;
begin
    {Инициализация}
    S := TTabSheet(Sender).Name;
    P := nil;
    If S = 'TS_Tab' then P := @GTab;
    If S = 'TS_Col' then P := @GCol;
    If S = 'TS_Row' then P := @GRow;
    If P = nil then Exit;
    ICaption := -1;
    IWidth   := P^.ClientWidth-1-GetSystemMetrics(SM_CXVSCROLL);

    {Просматриваем все столбцы}
    For I:=0 to P^.Columns.Count-1 do begin
       If P^.Columns.Items[I].FieldName <> F_CAPTION then IWidth:=IWidth-P^.Columns.Items[I].Width
                                                     else ICaption:=I;
    end;
    If ICaption = -1 then Exit;

    {Корректируем размер столбца-заголовка}
    P^.Columns.Items[ICaption].Width :=IWidth;

    {Корректируем размер кнопок}
    I := PBottom.ClientWidth Div 2 - 50;
    BtnOk.Width     := I;
    BtnCancel.Width := I;
end;


{==============================================================================}
{================  СОХРАНЕНИЕ/ВОССТАНОВЛЕНИЕ ШИРИНЫ СТОЛБЦОВ   ================}
{==============================================================================}
procedure TFKOD.SaveColWidth;
    procedure SaveGrid(const P: PDBGrid; const SIni: String);
    var Ind: Integer;
    begin
       For Ind:=0 to P^.Columns.Count-1 do begin
          WriteLocalInteger(INI_FORM_KOD,
                            SIni+P^.Columns.Items[Ind].FieldName,
                            P^.Columns.Items[Ind].Width);
       end;
    end;
begin
    SaveGrid(@GTab, FORM_KOD_TAB);
    SaveGrid(@GCol, FORM_KOD_COL);
    SaveGrid(@GRow, FORM_KOD_ROW);
end;

procedure TFKOD.LoadColWidth;
    procedure LoadGrid(const P: PDBGrid; const SIni: String);
    var Ind, IVal: Integer;
    begin
        For Ind:=0 to P^.Columns.Count-1 do begin
           IVal := ReadLocalInteger(INI_FORM_KOD, SIni+P^.Columns.Items[Ind].FieldName, -1);
           If IVal < 0 then begin
              IVal:=P^.Columns.Items[Ind].Width;
              If IVal > P^.ClientWidth then IVal:=P^.ClientWidth Div 3 * 2;
           end;
           P^.Columns.Items[Ind].Width := IVal;
        end;
    end;
begin
    LoadGrid(@GTab, FORM_KOD_TAB);
    LoadGrid(@GCol, FORM_KOD_COL);
    LoadGrid(@GRow, FORM_KOD_ROW);
end;

end.
