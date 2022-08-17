unit FunBD;

interface
uses
   Winapi.Windows,
   System.Classes, System.SysUtils,
   Vcl.Dialogs, Vcl.Controls,
   Data.DB, Data.Win.ADODB,
   FunType;

   // Windows - ��� SW_SHOWNORMAL

{��������� ��������� ������}
function  CloneQuery(const PQuery: PADOQuery): TADOQuery;
{��������� ��������� �������}
function  CloneTable(const PTable: PADOTable): TADOTable;
{�������� ���� ������ � ������}
function  OpenBD(const PBD       : PADOConnection;
                 const BDPath    : String;
                 const SPassword : String;
                 const PTables   : array of PADOTable;
                 const TabNames  : array of String): Boolean;
{�������� ���� ������ � ������}
procedure CloseBD(const PBD: PADOConnection; const PTables: array of PADOTable);
{������� ��� ������ �������}
procedure ClearTable(const PTable: PADOTable);
{������� ��� ������ ������ ������}
procedure ClearTableList(const PTables: array of PADOTable);
{������� ��� ����������� � �������}
procedure FreeTable(const PTable: PADOTable);
{������������� ������ �� �������}
procedure SetDBFilter(const PTable: PADOTable; const SFilter: String);
{������������� ���������� �� �������}
procedure SetDBSort(const PTable: PADOTable; const Sort: String);
{������������� ����� �� �������}
procedure SetDBConnect(const PTable: PADOTable;
                       const PMasterSource: PDataSource;
                       const MasterFields, IndexFields: String);
{���������� ����� ������� �������� ������� �� ������� ��������}
function  CountRecord_T(const PTable : PADOTable;
                        const FNames : array of String;
                        const VNames : array of String): Integer;
{����� ������ ���� ������}
function  FindBDTables(const PBD: PADOConnection; const LTableNames: array of String): Boolean;
{����� ������, ������������ � ���� ������}
function  FindDataSet(const PBD: PADOConnection; const TName: String): TADODataSet;
{���������� ������ (F_COUNTER) ������ �� ��������� �����}
function  GetIndRecord(const PBD: PADOConnection; const TName: String;
                       const FName, Val: array of String): String;
{������ �������� ���� ������ F_COUNTER}
function  ReadField(const PBD: PADOConnection; const TName: String; const ICounter: Integer;
                    const SField: String): String;
{������ �������� ���� SField ������, ������������ �������� SFilter}
function  ReadFieldFromFilter(const PBD: PADOConnection; const TName, SFilter, SField: String; const OnlyOneRec: Boolean): String;
{���������� �� ������, ������������ ��������: ���� ��, �� ������������� �� ��� ���������}
//function  IsExistsRec(const PData: PADODataSet; const LField, LVal: Array  of String): Boolean;


implementation

uses FunConst, FunSys, FunText, FunDay;


{==============================================================================}
{================ ��������� ��������� ������ ==================================}
{==============================================================================}
function CloneQuery(const PQuery: PADOQuery): TADOQuery;
begin
    Result:=TADOQuery(DuplicateComponents(TComponent(PQuery^)));
end;


{==============================================================================}
{================ ��������� ��������� ������� =================================}
{==============================================================================}
function CloneTable(const PTable: PADOTable): TADOTable;
begin
    Result:=TADOTable(DuplicateComponents(TComponent(PTable^)));
end;


{==============================================================================}
{====================== �������� ���� ������ � ������ =========================}
{==============================================================================}
function OpenBD(const PBD       : PADOConnection;
                const BDPath    : String;
                const SPassword : String;
                const PTables   : array of PADOTable;
                const TabNames  : array of String): Boolean;
var I: Integer;
begin
    Result:=false;
    If PBD=nil then Exit;
    // If Length(PTables)<>Length(TabNames) then Exit; ����� ������ ���� ���� �����

    {��������� ���� ������}
    If BDPath<>'' then begin
       {��������� ���������� ����}
       CloseBD(PBD, PTables);
       {���������� �� ����������� ���� ������}
       If FileExists(BDPath)=false then Exit;
       {������ �� ����������}
       PBD^.LoginPrompt := false;
       {��������� ���� ������}
       PBD^.ConnectionString := 'Provider=Microsoft.Jet.OLEDB.4.0; '+
                                'Data Source='+BDPath+'; '+
                                'Jet OLEDB:Database Password='+SPassword+'; '+
                                'Extended Properties=""';
    end;

    {��������� �������}
    For I:=Low(PTables) to High(PTables) do begin
       PTables[I]^.Connection := PBD^;
       If TabNames[I]<>'' then PTables[I]^.TableName := TabNames[I];
       PTables[I]^.Open;
    end;
    Result:=true;
end;


{==============================================================================}
{====================== �������� ���� ������ � ������ =========================}
{==============================================================================}
procedure CloseBD(const PBD: PADOConnection; const PTables: array of PADOTable);
var I: Integer;
begin
    {��������� �������}
    For I:=Low(PTables) to High(PTables) do PTables[I]^.Close;

    {��������� ���� ������}
    If PBD<>nil then begin
       If PBD^.Connected = true then PBD^.Close;
    end;
end;


{==============================================================================}
{======================  ������� ��� ������ �������   =========================}
{==============================================================================}
procedure ClearTable(const PTable: PADOTable);
begin
    FreeTable(PTable);
    PTable^.Last;
    While PTable^.Bof=false do PTable^.Delete;
end;


{==============================================================================}
{===============   ������� ��� ������ ������� ������ ������   =================}
{==============================================================================}
procedure ClearTableList(const PTables: array of PADOTable);
var I: Integer;
begin
    For I:=Low(PTables) to High(PTables) do ClearTable(PTables[I]);
end;


{==============================================================================}
{===================  ������� ��� ����������� � �������  ======================}
{==============================================================================}
procedure FreeTable(const PTable: PADOTable);
begin
    SetDBFilter (PTable, '');
    SetDBSort   (PTable, '');
    SetDBConnect(PTable, nil, '', '');
end;


{==============================================================================}
{====================  ������������� ������ �� �������  =======================}
{==============================================================================}
procedure SetDBFilter(const PTable: PADOTable; const SFilter: String);
begin
    With PTable^ do begin
       If (Filter = SFilter) and (Filtered = true) then Exit;

       If Filtered <> false then Filtered:=false;
       Filter := SFilter;
       If Filter <> '' then Filtered:=true;
    end;
end;


{==============================================================================}
{==================  ������������� ���������� �� ������� ======================}
{==============================================================================}
procedure SetDBSort(const PTable: PADOTable; const Sort: String);
begin
    // If PTable^.Active = false then Exit;
    If PTable^.Sort<>Sort then PTable^.Sort:=Sort;
end;


{==============================================================================}
{=====================  ������������� ����� �� �������  =======================}
{==============================================================================}
procedure SetDBConnect(const PTable                    : PADOTable;
                       const PMasterSource             : PDataSource;
                       const MasterFields, IndexFields : String);
var IsActive : Boolean;
begin
    IsActive                := PTable^.Active;
    If PTable^.Active then PTable^.Active := false;
    PTable^.IndexFieldNames := '';
    PTable^.MasterSource    := nil;
    If PMasterSource <> nil then PTable^.MasterSource := PMasterSource^
                            else PTable^.MasterSource := nil;
    PTable^.MasterFields    := MasterFields;
    PTable^.IndexFieldNames := IndexFields;         // ���������, �.�. ����� ��������������� ������������� ��� ��������� Master
    PTable^.Active          := IsActive;
//  !!! ������������ ����� � ����� �����������, ����� ��������� ������� ��� SQL
//    If PMasterSource<>nil then PTable^.MasterSource := PMasterSource^;
//    If PTable^.MasterFields    <> MasterFields   then PTable^.MasterFields    := MasterFields;
//    If PTable^.IndexFieldNames <> IndexFields    then PTable^.IndexFieldNames := IndexFields;
end;


{==============================================================================}
{=====  ���������� ����� ������� �������� ������� �� ������� ��������  ========}
{==============================================================================}
function CountRecord_T(const PTable : PADOTable;
                       const FNames : array of String;
                       const VNames : array of String): Integer;
var T0     : TADOTable;
    Filter : String;
    I      : Integer;
begin
    {�������������}
    Result:=0;
    If Length(FNames)<>Length(VNames) then Exit;
    If PTable<>nil then If PTable^.Active=false then Exit;
    Filter:='';

    {���������� �� ��������� � ������ ����}
    For I:=0 to Length(FNames)-1 do If PTable^.FindField(FNames[I])=nil then Exit;

    {���������� ������}
    For I:=0 to Length(FNames)-1 do Filter:=Filter+'(['+FNames[I]+']='+QuotedStr(VNames[I])+') AND ';
    Delete(Filter, Length(Filter)-3, 4);

    {������� ���� ������� ����� �� �������� �� ���������}
    T0:=CloneTable(PTable); If T0=nil then Exit;

    {���������� ����� ������� �������� �������}
    try
       SetDBFilter(@T0, Filter);
       Result:=T0.RecordCount;
    finally
       T0.Free;
    end;
end;



{==============================================================================}
{=======================   ����� ������ ���� ������   =========================}
{==============================================================================}
function FindBDTables(const PBD: PADOConnection; const LTableNames: array of String): Boolean;
var ITab  : Integer;
    SList : TStringList;
begin
    {�������������}
    Result := false;
    If PBD=nil then Exit;
    SList := TStringList.Create;
    try
       PBD^.GetTableNames(SList, false);
       Result := true;
       For ITab:=Low(LTableNames) to High(LTableNames) do begin
          Result := (SList.IndexOf(LTableNames[ITab]) >= 0);
          If Not Result then Break;
       end;
    finally
       SList.Free;
    end;
end;


{==============================================================================}
{===============   ����� ������, ������������ � ���� ������   =================}
{==============================================================================}
function FindDataSet(const PBD: PADOConnection; const TName: String): TADODataSet;
var I: Integer;
begin
    {�������������}
    Result:=nil;
    If PBD=nil then Exit;

    {������������� ������ ������}
    For I:=0 to PBD^.DataSetCount-1 do begin
       If CmpStr(TName, PBD^.DataSets[I].Name) then begin
          Result:=TADODataSet(PBD^.DataSets[I]);
          Break;
       end;
    end;
end;


{==============================================================================}
{=========  ���������� ������ (F_COUNTER) ������ �� ��������� �����  ==========}
{==============================================================================}
function GetIndRecord(const PBD: PADOConnection; const TName: String;
                      const FName, Val: array of String): String;
var SFilter : String;
    I       : Integer;
begin
    {�������������}
    Result:='';
    If PBD=nil then Exit;
    If TName='' then Exit;
    If Length(FName)<>Length(Val) then Exit;

    {������� ������}
    SFilter:='';
    For I:=Low(FName) to High(FName) do SFilter:=SFilter+' AND (['+FName[I]+']='+QuotedStr(Val[I])+')';
    Delete(SFilter, 1, 5);

    {���� � ������ ������}
    Result := ReadFieldFromFilter(PBD, TName, SFilter, F_COUNTER, true);
end;



{==============================================================================}
{==================  ������ �������� ���� ������ F_COUNTER   ==================}
{==============================================================================}
function ReadField(const PBD: PADOConnection; const TName: String; const ICounter: Integer; const SField: String): String;
begin
    Result := ReadFieldFromFilter(PBD, TName, '['+F_COUNTER+']='+IntToStr(ICounter), SField, true);
end;


{==============================================================================}
{====  ������ �������� ���� SFIELD ������, ������������ �������� SFILTER  =====}
{==============================================================================}
function ReadFieldFromFilter(const PBD: PADOConnection; const TName, SFilter, SField: String; const OnlyOneRec: Boolean): String;
var T: TADOTable;
begin
    {�������������}
    Result := '';
    T:=TADOTable.Create(nil);
    try
       {������������� �������}
       T.Connection := PBD^;
       T.TableName  := TName;
       T.Filter     := SFilter;
       T.Filtered   := true;
       T.Open;
       If OnlyOneRec then begin
          If T.RecordCount=1 then Result:=T.FieldByName(SField).AsString;
       end else begin
          If T.RecordCount>0 then begin
             T.First;
             Result:=T.FieldByName(SField).AsString;
          end;
       end;
    finally
       If T.Active then T.Close;
       T.Free;
    end;
end;

(*
{==============================================================================}
{==============   ���������� �� ������, ������������ ��������   ===============}
{==============           ������������� �� ��� ���������        ===============}
{==============================================================================}
function IsExistsRec(const PData: PADODataSet; const LField, LVal: array of String): Boolean;
var I, IHigh : Integer;
    FList    : array of TField;
begin
    {�������������}
    Result := false;
    If Not PData^.Active then Exit;
    IHigh  := High(LField);

    try
       SetLength(FList, Length(LField));
       For I:=0 to IHigh do FList[I]:=PData^.FindField(LField[I]);

       With PData^ do begin
          First;
          While Not Eof do begin
             For I:=0 to IHigh do begin
                If FList[I].AsString <> LVal[I] then Break;
                If I=IHigh then begin
                   Result:=true;
                   Exit;
                end;
             end;
             Next;
          end;
       end;
    finally
       SetLength(FList, 0);
    end;
end;
*)

end.

