{==============================================================================}
{===========    !!! ����� !!! ����� !!! ����� !!! ����� !!!     ===============}
{==============================================================================}
{===========  ������ � ����������� ������, YMR �� ������������  ===============}
{==============================================================================}
{===================   ������ ������ - ����� CashTab    =======================}
{===================   ������ ������ - ����� MemFile    =======================}
{==============================================================================}
unit MBASE;

interface

uses
   System.SysUtils, System.Variants, System.Classes, System.IniFiles,
   Vcl.Controls, Vcl.Forms,  Vcl.Dialogs,
   Data.DB, Data.Win.ADODB,
   MAIN, FunType;

type
   {��� ������}
   TCashTab = record
      SPath : String;       // ���� � ����������� dat-�����
      STab  : String;       // ���������� �������
      SList : TStringList;  // ��� �������
   end;

   TBASE = class
   public
      constructor Create;
      destructor  Destroy; override;
      procedure ClearCash;

      {������:  MBASE_READ}
      function  GetValYMR(const YMR, STab, SCol, SRow: String): Extended;
      function  GetVal(const SPath, STab, SCol, SRow: String): Extended;

      {������:  MBASE_WRITE}
      function  SetValYMR(const YMR, STab, SCol, SRow, SVal: String): Boolean;
      function  SetVal(const SPath, STab, SCol, SRow, SVal: String): Boolean;
      procedure SaveMemFile;

   private
      FFMAIN     : TFMAIN;
      CashTab    : array of TCashTab;
      MemFile    : array of TMemIniFile;
      BSEPDetail : Boolean;
      ISEPMonth  : Integer;

      {������:  MBASE_READ}
      function  FindCashDat(const SPath: String): Integer;
      function  FindCashTab(const SPath, STab: String): Integer;
      function  LoadCashDat(const SPath: String): Boolean;
      function  GetCash(const SPath, STab, SCol, SRow: String; var   IVal: Extended): Boolean;

      {������:  MBASE_WRITE}
      function  FindMemFile(const SPath: String): PMemIniFile;
      function  AddMemFile(const SPath: String): PMemIniFile;
   end;

implementation

uses FunConst, FunText, FunBD, FunExcel, FunSys, FunVcl, FunIO, FunDay, FunIni;

{$INCLUDE MBASE_WRITE}
{$INCLUDE MBASE_READ}

{==============================================================================}
{=============================   �����������   ================================}
{==============================================================================}
constructor TBASE.Create;
begin
    inherited Create;

    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    SetLength(CashTab, 0);
    SetLength(MemFile, 0);

    BSEPDetail := ReadLocalBool    (INI_SET, INI_SET_SEP_DETAIL, false);
    ISEPMonth  := ReadLocalInteger (INI_SET, INI_SET_SEP_MONTH,  18);
end;


{==============================================================================}
{==============================   ����������   ================================}
{==============================================================================}
destructor TBASE.Destroy;
begin
    ClearCash;

    inherited Destroy;
end;


{==============================================================================}
{=============================   ������� ����   ===============================}
{==============================================================================}
procedure TBASE.ClearCash;
var I : Integer;
begin
    For I:=Low(CashTab) to High(CashTab) do CashTab[I].SList.Free;
    SetLength(CashTab, 0);
    For I:=Low(MemFile) to High(MemFile) do MemFile[I].Free;
    SetLength(MemFile, 0);
end;

end.
