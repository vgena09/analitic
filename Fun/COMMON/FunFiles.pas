unit FunFiles;

interface

uses
   Winapi.Windows, Winapi.ShellApi, Winapi.ShlObj,
   Winapi.ActiveX,
   System.SysUtils, System.Classes,
   System.Win.Registry,
   Vcl.Forms;


{==============================================================================}

   function  GetPathWritable(const FPath: String): Boolean;

   function  GetFileSize(const FileName: String): LongInt;
   function  GetFileDate(const FileName: String): TDateTime;

   function  CreateEmptyFile(const SPath: String): Boolean;
   function  TempFileName(const Pref, Ext: String): String;
   function  CopyDir(const SDirSrc, SDirDect: String): Boolean;
   function  DelDir(const DirName: String; const DelRoot: Boolean): Boolean;

   function  WriteTxtFile(const SPath, Str: String): Boolean;
   function  ReadTxtFile(const SPath: String): String;

implementation


{==============================================================================}
{==================  ВОЗМОЖНОСТЬ ЗАПИСИ В УКАЗАННУЮ ПАПКУ   ===================}
{==============================================================================}
function GetPathWritable(const FPath: String): Boolean;
var F     : TextFile;
    SPath : String;
begin
    {Инициализация}
    Result := false;
    If FPath='' then Exit;
    If FPath[Length(FPath)]<>'\' then SPath := FPath+'\'
                                 else SPath := FPath;
    SPath := SPath+'Тест_удалить.$$$';

    {Проверка}
    try
       AssignFile(F, SPath);
       {$I-}
       Rewrite(F);
       If (IOResult = 0) then begin
          Result := true;
          CloseFile(F);
          DeleteFile(SPath);
       end;
       {$I+}
    finally
    end;
end;


{==============================================================================}
{======================  ОПРЕДЕЛИТЬ РАЗМЕР ФАЙЛА   ============================}
{==============================================================================}
function GetFileSize(const FileName: String): LongInt;
var InfoFile    : TSearchRec;
    AttrFile    : Integer;
    ErrorReturn : Integer;
begin
    Result   := 0;
    AttrFile := $0000003F; {Any file}
    try
       ErrorReturn := FindFirst(FileName, AttrFile, InfoFile);
       If ErrorReturn <> 0 then Result := -1             {в случае, если файл не найден}
                           else Result := InfoFile.Size; {Размер файла в байтах}
    finally
       FindClose(InfoFile);
    end;
end;


{==============================================================================}
{=================  ОПРЕДЕЛИТЬ ДАТУ/ВРЕМЯ ИЗМЕНЕНИЯ ФАЙЛА   ===================}
{==============================================================================}
function GetFileDate(const FileName: String): TDateTime;
var FHandle: Integer;
begin
    Result:=0;
    try FHandle := FileOpen(FileName, 0);
        try Result := FileDateToDateTime(FileGetDate(FHandle));
        finally FileClose(FHandle);
        end;
    finally
    end;
end;


{==============================================================================}
{========================   СОЗДАТЬ ПУСТОЙ ФАЙЛ   =============================}
{==============================================================================}
function CreateEmptyFile(const SPath: String): Boolean;
var hFile : Cardinal;
begin
    {Инициализация}
    Result:=false;
    If SPath='' then Exit;

    If Not FileExists(SPath) then begin
       {Создаем каталоги}
       If Not ForceDirectories(ExtractFilePath(SPath)) then Exit;

       {Создаем файл}
       try
          hFile := CreateFile(PChar(SPath), GENERIC_WRITE, 0, nil, OPEN_ALWAYS, 0, 0);
          CloseHandle(hFile);
       finally
       end;
    end;

    Result:=true;
end;


{==============================================================================}
{====================  СОЗДАТЬ УНИКАЛЬНОЕ ИМЯ ФАЙЛА   =========================}
{==============================================================================}
function TempFileName(const Pref, Ext: String): String;
var ICount: Integer;
begin
    Result:=Pref+'.'+Ext;
    If FileExists(Result)=false then Exit;

    ICount:=1;
    While true do begin
       Result:=Pref+IntToStr(ICount)+'.'+Ext;
       If FileExists(Result)=false then Break;
       Inc(ICount);
    end;
end;


{==============================================================================}
{========================  КОПИРОВАТЬ ДИРЕКТОРИЙ   ============================}
{==============================================================================}
function CopyDir(const SDirSrc, SDirDect: String): Boolean;
var fos: TSHFileOpStruct;
begin
    {Инициализация}
    Result:=false;
    If Not ForceDirectories(SDirDect) then Exit;

    ZeroMemory(@fos, SizeOf(fos));
    With fos do begin
       //Wnd    := Application.Handle;  просто отключил
       wFunc  := FO_COPY;
       fFlags := FOF_NOCONFIRMATION or FOF_SILENT;
       pFrom  := PChar(SDirSrc + #0);
       pTo    := PChar(SDirDect);
       fAnyOperationsAborted := false;
       hNameMappings         := nil;
       lpszProgressTitle     := nil;
    end;
    Result := (ShFileOperation(fos) = 0);
{   * FOF_ALLOWUNDO Если возможно, сохраняет информацию для возможности UnDo.
    * FOF_CONFIRMMOUSE Не реализовано.
    * FOF_FILESONLY Если в поле pFrom установлено *.*, то операция будет производиться только с файлами.
    * FOF_MULTIDESTFILES Указывает, что для каждого исходного файла в поле pFrom указана своя директория - адресат.
    * FOF_NOCONFIRMATION Отвечает "yes to all" на все запросы в ходе опеации.
    * FOF_NOCONFIRMMKDIR Не подтверждает создание нового каталога, если операция требует, чтобы он был создан.
    * FOF_RENAMEONCOLLISION В случае, если уже существует файл с данным именем, создается файл с именем "Copy #N of..."
    * FOF_SILENT Не показывать диалог с индикатором прогресса.
    * FOF_SIMPLEPROGRESS Показывать диалог с индикатором прогресса, но не показывать имен файлов.
    * FOF_WANTMAPPINGHANDLE Вносит hNameMappings элемент. Дескриптор должен быть освобожден функцией SHFreeNameMappings.
}
end;


{==============================================================================}
{====================  УДАЛИТЬ НЕПУСТОЙ ДИРЕКТОРИЙ   ==========================}
{==============================================================================}
{====================  DirName = C:\Temp\            ==========================}
{====================  DirName = C:\Temp             ==========================}
{==============================================================================}
function DelDir(const DirName: String; const DelRoot: Boolean): Boolean;
var FindData   : TWIN32FindData;
    FindHandle : THandle;
    SPath      : String;
begin
    Result := false;
    SPath  := Trim(DirName);
    If SPath='' then Exit;
    If SPath[Length(SPath)]<>'\' then SPath:=SPath+'\';
    try
       FindHandle := Winapi.Windows.FindFirstFile(PChar(SPath+'*.*'), FindData);
       While FindHandle <> INVALID_HANDLE_VALUE do begin
          If Not Winapi.Windows.FindNextFile(FindHandle, FindData) then Break;
          If (String(FindData.cFileName)='.') or (String(FindData.cFileName)='..') then Continue;
          If FindData.dwFileAttributes and FILE_ATTRIBUTE_DIRECTORY<>0 then begin
             DelDir(SPath+String(FindData.cFileName), true);
             RemoveDirectory(PChar(SPath+String(FindData.cFileName)));
          end;
          SetFileAttributes(PChar(SPath+String(FindData.cFileName)), FILE_ATTRIBUTE_NORMAL);
          DeleteFile(PChar(SPath+String(FindData.cFileName)));
       end;
       Winapi.Windows.FindClose(FindHandle);
       If DelRoot then RemoveDirectory(PChar(SPath));
       Result:=true;
   finally end;
end;


{==============================================================================}
{====================   ЗАПИСЬ В ТЕКСТОВЫЙ ФАЙЛ СТРОКИ    =====================}
{==============================================================================}
function WriteTxtFile(const SPath, Str: String): Boolean;
var F : TextFile;
begin
    AssignFile(F, SPath);
    try Rewrite(F);
        WriteLn(F, Str);
    finally
        CloseFile(F);
    end;
end;


{==============================================================================}
{====================   ЧТЕНИЕ СТРОКИ ТЕКСТОВОГО ФАЙЛА    =====================}
{==============================================================================}
function ReadTxtFile(const SPath: String): String;
var F : TextFile;
begin
    Result := '';
    AssignFile(F, SPath);
    try Reset(F);
        ReadLn(F, Result);
    finally
        CloseFile(F);
    end;
end;

end.

