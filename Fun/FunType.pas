unit FunType;

interface

{!!! Не должно быть El-компонентов, так как конфликт при выходе из DLL !!!}
uses System.Classes, System.IniFiles,
     Vcl.Forms, Vcl.ComCtrls, Vcl.Graphics, Vcl.Controls, Vcl.StdCtrls {TComboBox},
     Data.DB, Data.Win.ADODB;

type
   {*** Для операций с массивами в функциях ***********************************}
   TArrayStr = array of String;
   PArrayStr = ^TArrayStr;

   {Выбор в TreePath}
   TSelPath = record
      Node      : TTreeNode; // Выбранный узел
      SPath     : String;    // Выбранный путь: 2009\июнь\Республика Беларусь\Гродненская\Берестовицкий\
      IYear     : Integer;   // Выбранный год
      IMonth    : Integer;   // Выбранный месяц
      SYear     : String;    // Выбранный год
      SMonth    : String;    // Выбранный месяц
      SRegion   : String;    // Выбранный регион
      Date      : TDate;     // Дата смены периода
      DateBegin : TDate;     // Дата начала периода
      DateEnd   : TDate;     // Дата завершения периода
   end;
   PSelPath = ^TSelPath;

   {Выбор в TreeForm}
   TSelForm = record
      Node      : TTreeNode; // Выбранный узел
   end;
   PSelForm = ^TSelForm;

   {Параметры аналитики}
   TParam = record
      SCaption, SFormula, SColor : String;
   end;
   TListParam = array of TParam;
   PParam     = ^TParam;
   PListParam = ^TListParam;

   {Указатели}
   PBoolean       = ^Boolean;
   PString        = ^String;
   PStringList    = ^TStringList;
   PADOConnection = ^TADOConnection;
   PADOTable      = ^TADOTable;
   PADOQuery      = ^TADOQuery;
   PADODataSet    = ^TADODataSet;
   PDataSource    = ^TDataSource;
   //PButton        = ^TButton;
   PStrings       = ^TStrings;
   PTreeView      = ^TTreeView;
   PTreeNode      = ^TTreeNode;
   PIniFile       = ^TIniFile;
   PComboBox      = ^TComboBox;
   PMemIniFile    = ^TMemIniFile;
   PCustomIniFile = ^TCustomIniFile;
   PPBitmap       = ^TBitmap;          // !!! класс PBitmap в системе определен !!!

implementation

end.
