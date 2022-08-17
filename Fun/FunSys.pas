unit FunSys;

interface

uses
   Winapi.Windows, Winapi.Messages, Winapi.ShlObj, Winapi.ShellApi, Winapi.ActiveX,
   System.SysUtils, System.Variants, System.Classes, System.IniFiles,
   System.Win.Registry,
   Vcl.Forms, Vcl.Controls, Vcl.Graphics, Vcl.Dialogs, Vcl.ExtCtrls {TPanel},
   IdGlobal, IdGlobalProtocols,
   FunType;


type
   {Клонирование компонентов}
   TUniqueReader = Class(TReader)
      LastRead: TComponent;
      procedure ComponentRead(Component: TComponent);
      procedure SetNameUnique(Reader: TReader; Component: TComponent; var Name: string);
   end;

   {Глобальный поиск компонента}
   function  GlFindComponent(const CName: TComponentName): TComponent;
   function  FindChildComponent(var CParent: TComponent; const CName: TComponentName): TComponent;

   function  DuplicateComponents(AComponent: TComponent): TComponent;
   function  StrToBoolMy(const S: String): Boolean;
   function  BoolToStrMy(const B: Boolean): String;

   procedure StartAssociatedExe(const FPath: String; const ShowCmd: Integer);
   function  StartAssociatedAndWait(const FPath: String; const ShowCmd: Integer): Boolean;
   function  ExecAndWait(const FileName, Params: String; const WinState: Word): Boolean;

   function  GetKeyState(const vKey: Integer): Boolean;

   function  CopyFileToTemp(const SSrc: String): String;

implementation

{!!! Не должно быть El-компонентов, так как конфликт при выходе из DLL !!!}
uses FunConst, FunText, FunVcl, FunIni, FunFiles;


{==============================================================================}
{=====================  ГЛОБАЛЬНЫЙ  ПОИСК  КОМПОНЕНТА  ========================}
{==============================================================================}
function GlFindComponent(const CName: TComponentName): TComponent;
begin
   Result:=FindChildComponent(TComponent(Application), CName);
end;


{==============================================================================}
{=========================   ПОИСК  КОМПОНЕНТА   ==============================}
{==============================================================================}
function FindChildComponent(var CParent: TComponent; const CName: TComponentName): TComponent;
var I       : Integer;
    TComp   : TComponent;
    CNameUp : String;
begin
    {Инициализация}
    Result  := nil;
    CNameUp := AnsiUpperCase(CName);

    {Проверяем все компоненты}
    For I:=0 to CParent.ComponentCount-1 do begin

       {Проверяем имя субкомпонента}
       If AnsiUpperCase(CParent.Components[I].Name)=CNameUP then begin
          Result:=CParent.Components[I];
          Break;
       end;

       {Проверяем имена субсубкомпонентов}
       TComp:=CParent.Components[I];
       TComp:=FindChildComponent(TComp, CName);
       If TComp<>nil then begin
          Result:=TComp;
          Break;
       end;
    end;
end;


{==============================================================================}
{=============  К Л О Н И Р О В А Н И Е   К О М П О Н Е Н Т О В  ==============}
{==============================================================================}
procedure TUniqueReader.ComponentRead(Component: TComponent);
begin
   LastRead := Component;
end;

// Задаем уникальное имя считываемому компоненту, например, "Panel2", если "Panel1" уже существует
procedure TUniqueReader.SetNameUnique(Reader: TReader; Component: TComponent; var Name: string);
var i        : Integer;
    tempname : string;
begin
    i := 0;
    tempname := Name;
    while Component.Owner.FindComponent(Name) <> nil do begin
       Inc(i);
       Name := Format('%s%d', [tempname, i]);
    end;
end;

function DuplicateComponents(AComponent: TComponent): TComponent;
     procedure RegisterComponentClasses(AComponent: TComponent);
     var i : integer;   
     begin
         RegisterClass(TPersistentClass(AComponent.ClassType));
         if AComponent is TWinControl then
            if TWinControl(AComponent).ControlCount > 0 then
               for i := 0 to (TWinControl(AComponent).ControlCount-1) do
                   RegisterComponentClasses(TWinControl(AComponent).Controls[i]);
     end;
var Stream       : TMemoryStream;
    UniqueReader : TUniqueReader;
    Writer       : TWriter;
begin
    result := nil;
    UniqueReader := nil;
    Writer := nil;
    Stream := TMemoryStream.Create;
    try
       RegisterComponentClasses(AComponent);
       try
          Writer := TWriter.Create(Stream, 4096);
          Writer.Root := AComponent.Owner;
          Writer.WriteSignature;
          Writer.WriteComponent(AComponent);
          Writer.WriteListEnd;
       finally
          Writer.Free;
       end;
       Stream.Position := 0;
       try
          UniqueReader := TUniqueReader.Create(Stream, 4096);     // создаем поток, перемещающий данные о компоненте в конструктор
          UniqueReader.OnSetName := UniqueReader.SetNameUnique;
          UniqueReader.LastRead := nil;
          if AComponent is TWinControl then
               UniqueReader.ReadComponents(TWinControl(AComponent).Owner,
                                           TWinControl(AComponent).Parent,
                                           UniqueReader.ComponentRead)
          else UniqueReader.ReadComponents(AComponent.Owner,
                                           nil,
                                           UniqueReader.ComponentRead);
          result := UniqueReader.LastRead;
       finally
          UniqueReader.Free;
       end;
    finally
       Stream.Free;
    end;
end;


{==============================================================================}
{====================    ПЕРЕВОД  STRING  В  BOOLEAN     ======================}
{==============================================================================}
function StrToBoolMy(const S: String): Boolean;
begin
    Result:=FindStrInArray(S, ['TRUE', 'ИСТИНА', 'ДА']);
end;


{==============================================================================}
{====================    ПЕРЕВОД  BOOLEAN  В  STRING     ======================}
{==============================================================================}
function BoolToStrMy(const B: Boolean): String;
begin
    If B then Result:='true' else Result:='false';
end;


{==============================================================================}
{=======================  ЗАПУСК ВНЕШНЕГО ПРИЛОЖЕНИЯ   ========================}
{==============================================================================}
{============  ShowCmd - SW_SHOWMAXIMIZED, SW_SHOWNORMAL, ...   ===============}
{==============================================================================}
procedure StartAssociatedExe(const FPath: String; const ShowCmd: Integer);
begin
    ShellExecute(Application.Handle, 'open', PChar(FPath), nil, nil, ShowCmd);
end;


{==============================================================================}
{=======  ЗАПУСК АССОЦИИРОВАННОГО ПРИЛОЖЕНИЯ И ОЖИДАНИЕ ЕГО ЗАВЕРШЕНИЯ  =======}
{==============================================================================}
{============  ShowCmd - SW_SHOWMAXIMIZED, SW_SHOWNORMAL, ...   ===============}
{==============================================================================}
function StartAssociatedAndWait(const FPath: String; const ShowCmd: Integer): Boolean;
var SEInfo : TShellExecuteInfo;
    ExCode : DWORD;
begin
    {Инициализация}
    Result:=false;
    FillChar(SEInfo, SizeOf(SEInfo), 0);
    SEInfo.cbSize := SizeOf(TShellExecuteInfo);
    With SEInfo do begin
       fMask  := SEE_MASK_NOCLOSEPROCESS;
       Wnd    := Application.Handle;
       lpFile := PChar(FPath);
       nShow  := ShowCmd;
    end;

    {Запускаем приложение}
    If Not ShellExecuteEx(@SEInfo) then Exit;

    {Ожидаем завершение приложения}
    repeat GetExitCodeProcess(SEInfo.hProcess, ExCode);
    until (ExCode <> STILL_ACTIVE) or Application.Terminated;

    {Возвращаемый результат}
    Result:=true;
end;


{==============================================================================}
{=====      ЗАПУСК ВНЕШНЕГО ПРИЛОЖЕНИЯ И ОЖИДАНИЕ ЕГО ЗАВЕРШЕНИЯ         ======}
{==============================================================================}
{===== WinState - SW_HIDE, SW_MAXIMIZE, SW_MINIMIZE, SW_SHOWNORMAL       ======}
{==============================================================================}
function ExecAndWait(const FileName, Params: String; const WinState: Word): Boolean;
var StartInfo : TStartupInfo;
    ProcInfo  : TProcessInformation;
    CmdLine   : String;
begin
    {Помещаем имя файла между кавычками, с соблюдением всех пробелов в именах Win9x }
    CmdLine := '"' + Filename + '" ' + Params;
    FillChar(StartInfo, SizeOf(StartInfo), #0);
    With StartInfo do begin
      cb          := SizeOf(StartUPInfo); //(SUInfo);
      dwFlags     := STARTF_USESHOWWINDOW;
      wShowWindow := WinState;
    end;
    Result := CreateProcess(nil, PChar(String(CmdLine)), nil, nil, false,
                            CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil,
                            PChar(ExtractFilePath(Filename)), StartInfo, ProcInfo);

    {Ожидаем завершения приложения}
    If Result then begin
       WaitForSingleObject(ProcInfo.hProcess, INFINITE);
       CloseHandle(ProcInfo.hProcess);
       CloseHandle(ProcInfo.hThread);
  end;
end;


{==============================================================================}
{==========================  НАЖАТА ЛИ КЛАВИША  ===============================}
{==============================================================================}
function GetKeyState(const vKey: Integer): Boolean;
begin
    Result := GetAsyncKeyState(vKey) <> 0;    // VK_LCONTROL
end;


{==============================================================================}
{=============  СОЗДАНИЕ КОПИИ ФАЙЛА SSrc ВО ВРЕМЕННЫЙ КАТАЛОГ  ===============}
{==============================================================================}
function CopyFileToTemp(const SSrc: String): String;
var SDect : String;
    SExt  : String;
begin
    {Инициализация}
    Result:='';

    {Создаем временную копию}
    SExt  := ExtractFileExt(SSrc);
    Delete(SExt, 1, 1);
    SDect := PATH_WORK_TEMP+TEMP_XLS;
    SDect := TempFileName(ExtractFilePath(SDect)+ExtractFileNameWithoutExt(SDect), SExt);
    If Not ForceDirectories(ExtractFilePath(SDect)) then begin ErrMsg('Ошибка создания каталога: '+ExtractFilePath(SDect)); Exit; end;
    If Not CopyFileTo(SSrc, SDect)                  then begin ErrMsg('Ошибка создания файла отчета: '+SDect); Exit; end;

    {Возвращаемый результат}
    Result:=SDect;
end;

end.

