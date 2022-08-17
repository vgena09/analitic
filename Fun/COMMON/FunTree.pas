unit FunTree;

interface

uses Winapi.Windows {TreeView_SetItem}, Winapi.CommCtrl,
     System.SysUtils, System.Classes, System.IniFiles,
     Vcl.Dialogs, Vcl.ComCtrls,
     FunType;

{������������� �������� ������ � ��������� ������}
procedure SetNodeState(const PNode: PTreeNode; const Flags: Integer);
{��������� ���� ����}
function  GetNodePath(const PNode: PTreeNode): String;
{���� ������ ���� � ��������� SPath}
function  FindNodePath(const PTree: PTreeView; const SPath: String; const IsAll: Boolean): TTreeNode;
{���� ������ ���� � ���������� SText � PData}
function  FindNode(const PTree: PTreeView; const SText: String; const PData: Pointer): TTreeNode;
{���� ������ ���� � ��������� SText}
function  FindNodeText(const PTree: PTreeView; const SText: String): TTreeNode;
{���� ������ ���� � ��������� PData}
function  FindNodeData(const PTree: PTreeView; const PData: Pointer): TTreeNode;
{��������� ��������� � Tree}
procedure SaveTreeSelect(const PTree: PTreeView; const Section, Key: String);
{��������������� ��������� � Tree}
procedure LoadTreeSelect(const PTree: PTreeView; const Section, Key: String);
{������������� Tree � ����������� �� ���������}
procedure AutoScrollTree(const PTree: PTreeView);

implementation

uses FunConst, FunText;


{==============================================================================}
{=====         ������������� �������� ������ � ��������� ������       =========}
{=====   TVIS_BOLD - ����� ��p��� or TVIS_CUT-������� ������ 0-�����  =========}
{==============================================================================}
procedure SetNodeState(const PNode: PTreeNode; const Flags: Integer);
var tvi: TTVItem;
begin
    If PNode=nil then Exit;
    If not Assigned(PNode^) then Exit;
    FillChar(tvi, Sizeof(tvi), 0);
    tvi.hItem := PNode^.ItemID;
    tvi.mask := TVIF_STATE;
    tvi.stateMask := TVIS_BOLD or TVIS_CUT;
    tvi.state := Flags;
    TreeView_SetItem(PNode^.Handle, tvi);
end;


{==============================================================================}
{=========================   ��������� ���� ����   ============================}
{==============================================================================}
function GetNodePath(const PNode: PTreeNode): String;
var N: TTreeNode;
begin
    {�������������}
    Result:='';
    N:=PNode^;

    {������������� ����}
    While N<>nil do begin
       Result := N.Text+'\'+Result;
       N      := N.Parent;
    end;
end;


{==============================================================================}
{=================     ���� ������ ���� � ��������� SPATH     =================}
{==============================================================================}
function FindNodePath(const PTree: PTreeView; const SPath: String; const IsAll: Boolean): TTreeNode;
var N, Result0 : TTreeNode;
    S          : String;
    I, ILength : Integer;
begin
    {�������������}
    Result  := nil;
    Result0 := nil;
    ILength := 0;
    If PTree=nil then Exit;

    {������������� ����}
    N:=PTree^.Items.GetFirstNode;
    While N<>nil do begin
        S:=GetNodePath(@N);

        {������ ����������}
        If S=SPath then begin
           Result:=N;
           Break;
        end;

        {��������� ����������}
        If Not IsAll then begin
           I:=CmpStrFirst(S, SPath);
           If I>ILength then begin ILength:=I; Result0:=N; end;
        end;

        {��������� ����}
        N:=N.GetNext;
    end;

    {��� ��������� ����������}
    If (Result=nil) and (not IsAll) and (Result0<>nil) then Result:=Result0;
end;


{==============================================================================}
{=============    ���� ������ ���� � ���������� STEXT � PDATA    ==============}
{==============================================================================}
function FindNode(const PTree: PTreeView; const SText: String; const PData: Pointer): TTreeNode;
var N: TTreeNode;
begin
    {�������������}
    Result:=nil;
    If PTree=nil then Exit;

    N:=PTree^.Items.GetFirstNode;
    While N<>nil do begin
        If (N.Data=PData) and (N.Text=SText) then begin
           Result:=N;
           Break;
        end;
        N:=N.GetNext;
    end;
end;


{==============================================================================}
{=================     ���� ������ ���� � ��������� STEXT     =================}
{==============================================================================}
function FindNodeText(const PTree: PTreeView; const SText: String): TTreeNode;
var N: TTreeNode;
begin
    {�������������}
    Result:=nil;
    If PTree=nil then Exit;

    N:=PTree^.Items.GetFirstNode;
    While N<>nil do begin
        If N.Text=SText then begin
           Result:=N;
           Break;
        end;
        N:=N.GetNext;
    end;
end;


{==============================================================================}
{=================     ���� ������ ���� � ��������� PDATA     =================}
{==============================================================================}
function FindNodeData(const PTree: PTreeView; const PData: Pointer): TTreeNode;
var N: TTreeNode;
begin
    {�������������}
    Result:=nil;
    If PTree=nil then Exit;

    N:=PTree^.Items.GetFirstNode;
    While N<>nil do begin
        If N.Data=PData then begin
           Result:=N;
           Break;
        end;
        N:=N.GetNext;
    end;
end;


{==============================================================================}
{======================   ��������� ��������� � TREE    =======================}
{==============================================================================}
procedure SaveTreeSelect(const PTree: PTreeView; const Section, Key: String);
var F : TIniFile;
    N : TTreeNode;
begin
    {���������� ���������}
    N:=PTree^.Selected;
    F:=TIniFile.Create(PATH_WORK_INI);
    try     F.WriteString(Section, Key, GetNodePath(@N));
    finally F.Free; end;
end;


{==============================================================================}
{==================   ��������������� ��������� � TREE    =====================}
{==============================================================================}
procedure LoadTreeSelect(const PTree: PTreeView; const Section, Key: String);
var Event : TTVChangedEvent;
    F     : TIniFile;
    N     : TTreeNode;
    SPath : String;
begin
    If PTree=nil then Exit;

    {������ ���� ���������}
    F:=TIniFile.Create(PATH_WORK_INI);
    try     SPath := F.ReadString(Section, Key, '');
    finally F.Free; end;
    If SPath = '' then Exit;

    {���� ���� ���������}
    N := FindNodePath(PTree, SPath, false);
    If N=nil then begin
       If PTree^.Items.Count=0 then Exit;
       N:=PTree^.Items[0]; 
    end;

    {������������� ���������}
    Event           := PTree^.OnChange;
    PTree^.OnChange := nil;
    PTree.Selected  := N;
    PTree^.OnChange := Event;
end;


{==============================================================================}
{===========   ������������� TREE � ����������� �� ���������    ===============}
{==============================================================================}
procedure AutoScrollTree(const PTree: PTreeView);
var N1, N2 : TTreeNode;
    I      : Integer;
begin
    If PTree = nil then Exit;
    If Not Assigned(PTree^.Selected) then Exit;
    N1 := PTree^.Selected;
    N2 := N1;
    For I := 1 to (PTree^.Height Div 30) do begin
       If N1 <> nil then begin
          N2 := N1;
          N1 := N1.GetPrevVisible;
       end;
    end;   
    PTree^.TopItem := N2;
end;

end.

