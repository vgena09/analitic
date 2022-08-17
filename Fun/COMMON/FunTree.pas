unit FunTree;

interface

uses Winapi.Windows {TreeView_SetItem}, Winapi.CommCtrl,
     System.SysUtils, System.Classes, System.IniFiles,
     Vcl.Dialogs, Vcl.ComCtrls,
     FunType;

{Устанавливает жирность шрифта и бледность иконки}
procedure SetNodeState(const PNode: PTreeNode; const Flags: Integer);
{Формирует путь узла}
function  GetNodePath(const PNode: PTreeNode): String;
{Ищет первый узел с указанным SPath}
function  FindNodePath(const PTree: PTreeView; const SPath: String; const IsAll: Boolean): TTreeNode;
{Ищет первый узел с указанными SText и PData}
function  FindNode(const PTree: PTreeView; const SText: String; const PData: Pointer): TTreeNode;
{Ищет первый узел с указанным SText}
function  FindNodeText(const PTree: PTreeView; const SText: String): TTreeNode;
{Ищет первый узел с указанным PData}
function  FindNodeData(const PTree: PTreeView; const PData: Pointer): TTreeNode;
{Сохраняет выделение в Tree}
procedure SaveTreeSelect(const PTree: PTreeView; const Section, Key: String);
{Восстанавливает выделение в Tree}
procedure LoadTreeSelect(const PTree: PTreeView; const Section, Key: String);
{Автопрокрутка Tree в зависимости от выделения}
procedure AutoScrollTree(const PTree: PTreeView);

implementation

uses FunConst, FunText;


{==============================================================================}
{=====         Устанавливает жирность шрифта и бледность иконки       =========}
{=====   TVIS_BOLD - текст жиpным or TVIS_CUT-бледная иконка 0-сброс  =========}
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
{=========================   ФОРМИРУЕТ ПУТЬ УЗЛА   ============================}
{==============================================================================}
function GetNodePath(const PNode: PTreeNode): String;
var N: TTreeNode;
begin
    {Инициализация}
    Result:='';
    N:=PNode^;

    {Просматриваем узлы}
    While N<>nil do begin
       Result := N.Text+'\'+Result;
       N      := N.Parent;
    end;
end;


{==============================================================================}
{=================     ИЩЕТ ПЕРВЫЙ УЗЕЛ С УКАЗАННЫМ SPATH     =================}
{==============================================================================}
function FindNodePath(const PTree: PTreeView; const SPath: String; const IsAll: Boolean): TTreeNode;
var N, Result0 : TTreeNode;
    S          : String;
    I, ILength : Integer;
begin
    {Инициализация}
    Result  := nil;
    Result0 := nil;
    ILength := 0;
    If PTree=nil then Exit;

    {Просматриваем узлы}
    N:=PTree^.Items.GetFirstNode;
    While N<>nil do begin
        S:=GetNodePath(@N);

        {Полное совпадение}
        If S=SPath then begin
           Result:=N;
           Break;
        end;

        {Частичное совпадение}
        If Not IsAll then begin
           I:=CmpStrFirst(S, SPath);
           If I>ILength then begin ILength:=I; Result0:=N; end;
        end;

        {Следующий узел}
        N:=N.GetNext;
    end;

    {При частичном совпадении}
    If (Result=nil) and (not IsAll) and (Result0<>nil) then Result:=Result0;
end;


{==============================================================================}
{=============    ИЩЕТ ПЕРВЫЙ УЗЕЛ С УКАЗАННЫМИ STEXT И PDATA    ==============}
{==============================================================================}
function FindNode(const PTree: PTreeView; const SText: String; const PData: Pointer): TTreeNode;
var N: TTreeNode;
begin
    {Инициализация}
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
{=================     ИЩЕТ ПЕРВЫЙ УЗЕЛ С УКАЗАННЫМ STEXT     =================}
{==============================================================================}
function FindNodeText(const PTree: PTreeView; const SText: String): TTreeNode;
var N: TTreeNode;
begin
    {Инициализация}
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
{=================     ИЩЕТ ПЕРВЫЙ УЗЕЛ С УКАЗАННЫМ PDATA     =================}
{==============================================================================}
function FindNodeData(const PTree: PTreeView; const PData: Pointer): TTreeNode;
var N: TTreeNode;
begin
    {Инициализация}
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
{======================   СОХРАНЯЕТ ВЫДЕЛЕНИЕ В TREE    =======================}
{==============================================================================}
procedure SaveTreeSelect(const PTree: PTreeView; const Section, Key: String);
var F : TIniFile;
    N : TTreeNode;
begin
    {Запоминаем выделение}
    N:=PTree^.Selected;
    F:=TIniFile.Create(PATH_WORK_INI);
    try     F.WriteString(Section, Key, GetNodePath(@N));
    finally F.Free; end;
end;


{==============================================================================}
{==================   ВОССТАНАВЛИВАЕТ ВЫДЕЛЕНИЕ В TREE    =====================}
{==============================================================================}
procedure LoadTreeSelect(const PTree: PTreeView; const Section, Key: String);
var Event : TTVChangedEvent;
    F     : TIniFile;
    N     : TTreeNode;
    SPath : String;
begin
    If PTree=nil then Exit;

    {Читаем путь выделения}
    F:=TIniFile.Create(PATH_WORK_INI);
    try     SPath := F.ReadString(Section, Key, '');
    finally F.Free; end;
    If SPath = '' then Exit;

    {Ищем узел выделения}
    N := FindNodePath(PTree, SPath, false);
    If N=nil then begin
       If PTree^.Items.Count=0 then Exit;
       N:=PTree^.Items[0]; 
    end;

    {Устанавливаем выделение}
    Event           := PTree^.OnChange;
    PTree^.OnChange := nil;
    PTree.Selected  := N;
    PTree^.OnChange := Event;
end;


{==============================================================================}
{===========   АВТОПРОКРУТКА TREE В ЗАВИСИМОСТИ ОТ ВЫДЕЛЕНИЯ    ===============}
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

