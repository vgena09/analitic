unit FunType;

interface

{!!! �� ������ ���� El-�����������, ��� ��� �������� ��� ������ �� DLL !!!}
uses System.Classes, System.IniFiles,
     Vcl.Forms, Vcl.ComCtrls, Vcl.Graphics, Vcl.Controls, Vcl.StdCtrls {TComboBox},
     Data.DB, Data.Win.ADODB;

type
   {*** ��� �������� � ��������� � �������� ***********************************}
   TArrayStr = array of String;
   PArrayStr = ^TArrayStr;

   {����� � TreePath}
   TSelPath = record
      Node      : TTreeNode; // ��������� ����
      SPath     : String;    // ��������� ����: 2009\����\���������� ��������\�����������\�������������\
      IYear     : Integer;   // ��������� ���
      IMonth    : Integer;   // ��������� �����
      SYear     : String;    // ��������� ���
      SMonth    : String;    // ��������� �����
      SRegion   : String;    // ��������� ������
      Date      : TDate;     // ���� ����� �������
      DateBegin : TDate;     // ���� ������ �������
      DateEnd   : TDate;     // ���� ���������� �������
   end;
   PSelPath = ^TSelPath;

   {����� � TreeForm}
   TSelForm = record
      Node      : TTreeNode; // ��������� ����
   end;
   PSelForm = ^TSelForm;

   {��������� ���������}
   TParam = record
      SCaption, SFormula, SColor : String;
   end;
   TListParam = array of TParam;
   PParam     = ^TParam;
   PListParam = ^TListParam;

   {���������}
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
   PPBitmap       = ^TBitmap;          // !!! ����� PBitmap � ������� ��������� !!!

implementation

end.
