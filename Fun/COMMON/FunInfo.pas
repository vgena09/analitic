unit FunInfo;

interface

uses
   Winapi.Windows, Winapi.ShellApi, Winapi.ShlObj,
   Winapi.ActiveX,
   System.SysUtils,
   System.Win.Registry,
   Vcl.Forms;

type
   {Версия Windows}
   TWinVer   = (wvUnknown, wv95, wv98, wvME, wvNT3, wvNT4, wvW2K, wvXP, wv2003, wvVista, wv7);
   {Версия MSOffise}
   TMSOffice = (msWord, msExcel, msAccess, msPowerPoint, msOutlook);
   {Информация о программе}
   TVersionInfo = record
      CompanyName: WideString;
      FileDescription: WideString;
      FileVersion: WideString;
      InternalName: WideString;
      LegalCopyright: WideString;
      LegalTradeMarks: WideString;
      OriginalFilename: WideString;
      ProductName: WideString;
      ProductVersion: WideString;
      Comments: WideString;
      Language: WideString;
      Translation: WideString;

      FileVersionMajor: Word;
      FileVersionMinor: Word;
      FileVersionRelease: Word;
      FileVersionBuild: Word;
      ProductVersionMajor: Word;
      ProductVersionMinor: Word;
      ProductVersionRelease: Word;
      ProductVersionBuild: Word;
   end;
   {Ассоциация расширений}
   TAssocStr = (ASSOCSTR_COMMAND = 1,
                ASSOCSTR_EXECUTABLE,
                ASSOCSTR_FRIENDLYDOCNAME,
                ASSOCSTR_FRIENDLYAPPNAME,
                ASSOCSTR_NOOPEN,
                ASSOCSTR_SHELLNEWVALUE,
                ASSOCSTR_DDECOMMAND,
                ASSOCSTR_DDEIFEXEC,
                ASSOCSTR_DDEAPPLICATION,
                ASSOCSTR_DDETOPIC);

   function AssocQueryString(Flags: Integer; StrType: TAssocStr;
                             pszAssoc, pszExtra: PChar; pszOut: PChar;
                             var pcchPut: DWORD): HRESULT; stdcall;
                             external 'shlwapi.dll' name 'AssocQueryStringA';

{==============================================================================}

   function  GetWinUser: String;
   function  GetWinVer: TWinVer;
   function  GetWinVerStr: String;
   function  GetMSOfficeVer(const msProgram: TMSOffice): Integer;

   function  GetPathMyDoc: String;

   function  GetProgramAssociation(const Ext: String): String;
   function  GetProgRun: Boolean;



   function  GetProgCompanyName: String;
   function  GetProgFileDescription: String;
   function  GetProgFileVersion: String;
   function  GetProgInternalName: String;
   function  GetProgLegalCopyright: String;
   function  GetProgLegalTradeMarks: String;
   function  GetProgOriginalFilename: String;
   function  GetProgProductName: String;
   function  GetProgProductVersion: String;
   function  GetProgComments: String;
   function  GetProgLanguage: String;
   function  GetProgTranslation: String;
   function  GetProgInfo(const FileName: String; var VersionInfo: TVersionInfo): Boolean;

const
   WinVerStr : array[TWinVer] of string = ('Windows ?',   'Windows 95', 'Windows 98',
             'Windows ME', 'Windows NT 3', 'Windows NT 4',  'Windows 2000',
             'Windows XP', 'Windows 2003', 'Windows Vista', 'Windows 7');

implementation


{==============================================================================}
{==========================    ИМЯ ПОЛЬЗОВАТЕЛЯ    ============================}
{==============================================================================}
function GetWinUser: String;
var UserName    : String;
    UserNameLen : Dword;
begin
    UserNameLen := 255;
    SetLength(UserName, UserNameLen);
    If GetUserName(PChar(UserName), UserNameLen)
    then Result := Copy(UserName, 1, UserNameLen - 1)
    else Result := 'Не определен';
end;


{==============================================================================}
{===========================    ВЕРСИЯ WINDOWS    =============================}
{==============================================================================}
function GetWinVer: TWinVer;
var OSVersionInfo : TOSVersionInfo;
begin
    Result := wvUnknown;           // Неизвестная версия ОС
    OSVersionInfo.dwOSVersionInfoSize := sizeof(TOSVersionInfo);
    If GetVersionEx(OSVersionInfo) then begin
       Case OSVersionInfo.DwMajorVersion of
       3: Result := wvNT3;         // Windows NT 3
       4: Case OSVersionInfo.DwMinorVersion of
          0:  if OSVersionInfo.dwPlatformId = VER_PLATFORM_WIN32_NT
              then Result := wvNT4 // Windows NT 4
              else Result := wv95; // Windows 95
          10: Result := wv98;      // Windows 98
          90: Result := wvME;      // Windows ME
          end;
       5: Case OSVersionInfo.DwMinorVersion of
          0:  Result := wvW2K;     // Windows 2000
          1:  Result := wvXP;      // Windows XP
          2:  Result := wv2003;    // Windows 2003
          end;
       6: Case OSVersionInfo.DwMinorVersion of
          0:  Result := wvVista;   // Windows Vista
          1:  Result := wv7;       // Windows 7
          end;
       end;
    end;
end;


{==============================================================================}
{====================    ВЕРСИЯ WINDOWS: НАИМЕНОВАНИЕ    ======================}
{==============================================================================}
function GetWinVerStr: String;
begin
    Result := WinVerStr[GetWinVer];
end;


{==============================================================================}
{======================   ВЕРСИЯ MICROSOFT OFFICE   ===========================}
{==============================================================================}
function GetMSOfficeVer(const msProgram: TMSOffice): Integer;
const Key = 'SOFTWARE\MICROSOFT\OFFICE\%D.0\%S';
var Reg    : TRegistry;
    prName : String;
    I      : Integer;
begin
    {Инициализация}
    Result := 0;
    case msProgram of
       msWord       : prName := 'WORD';
       msExcel      : prName := 'EXCEL';
       msAccess     : prName := 'ACCESS';
       msPowerPoint : prName := 'POWERPOINT';
       msOutlook    : prName := 'OUTLOOK';
    end;

    {Читаем реестр}
    Reg := TRegistry.Create;
    try
       Reg.Rootkey := HKEY_LOCAL_MACHINE;
       For I:=1 to 50 do begin // С запасом
          If Reg.OpenKey(Format(Key, [I, prName]), false) then begin
             Reg.CloseKey;
             Result := I;
             Break;
          end;
       end;
  finally
      Reg.Free;
  end;
end;


{==============================================================================}
{=================  РАСПОЛОЖЕНИЕ ПАПКИ "МОИ ДОКУМЕНТЫ"   ======================}
{==============================================================================}
function GetPathMyDoc: String;
var Buf  : array [0..MAX_PATH] of Char;
    PIDL : PItemIDList;
begin
    Result := '';
    SHGetSpecialFolderLocation(0, CSIDL_PERSONAL, PIDL);
    try
       FillChar(Buf, SizeOf(Buf), 0);
       If SHGetPathFromIDList(PIDL, @Buf[0]) then Result := PChar(@Buf[0]);
    finally
       CoTaskMemFree(PIDL);
    end;
end;


{==============================================================================}
{============  ПРОГРАММА, С КОТОРОЙ АССОЦИИРОВАНО РАСШИРЕНИЕ  =================}
{==============================================================================}
{============                   SExt - '.ext'                 =================}
{==============================================================================}
// При возможности переделать на обращение к функции, а не прямое обращение к реестру
function GetProgramAssociation(const Ext: String): String;
var sExtDesc: string;
begin
    With TRegistry.Create do begin
       try
          RootKey:= HKEY_CLASSES_ROOT;
          If OpenKeyReadOnly(Ext) then begin
             sExtDesc:= ReadString('');
             CloseKey;
          end;
          If sExtDesc  <> '' then begin
             If OpenKeyReadOnly(sExtDesc + '\Shell\Open\Command') then begin
                Result:= ReadString('');
             end
          end;
      finally
        Free;
      end;
    end;

    If Result <> '' then begin
       If Result[1] = '"' then begin
          Result:= Copy(Result, 2, -1 + Pos('"', Copy(Result, 2, MaxINt))) ;
       end
    end;
end;


{==============================================================================}
{==========================   ЗАПУЩЕНА ЛИ ПРОГРАММА   =========================}
{==============================================================================}
function GetProgRun: Boolean;
var hMutex: THandle;
begin
    {Инициализация}
    Result:=false;
    try
       hMutex := CreateMutex(nil, true , PWideChar(GetProgProductName));
       If GetLastError = ERROR_ALREADY_EXISTS then begin
          CloseHandle(hMutex);
          Result:=true;
       end;
    finally
    end;
end;


{==============================================================================}
{==========================  ИНФОРМАЦИЯ О ПРОГРАММЕ  ==========================}
{==============================================================================}
function GetProgCompanyName: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.CompanyName else Result :='';
end;
function GetProgFileDescription: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.FileDescription else Result :='';
end;
function GetProgFileVersion: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.FileVersion else Result :='';
end;
function GetProgInternalName: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.InternalName else Result :='';
end;
function GetProgLegalCopyright: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.LegalCopyright else Result :='';
end;
function GetProgLegalTradeMarks: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.LegalTradeMarks else Result :='';
end;
function GetProgOriginalFilename: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.OriginalFilename else Result :='';
end;
function GetProgProductName: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.ProductName else Result :='';
end;
function GetProgProductVersion: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.ProductVersion else Result :='';
end;
function GetProgComments: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.Comments else Result :='';
end;
function GetProgLanguage: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.Language else Result :='';
end;
function GetProgTranslation: String;
var VersionInfo: TVersionInfo;
begin
    If GetProgInfo(Application.EXEName, VersionInfo) then Result := VersionInfo.Translation else Result :='';
end;

function GetProgInfo(const FileName: String; var VersionInfo: TVersionInfo): Boolean;
var Handle, Len, Size : Cardinal;
    Translation       : WideString;
    Data              : PWideChar;
    Buffer            : Pointer;
    FixedFileInfo     : PVSFixedFileInfo;
begin
    Result := False;
    With VersionInfo do begin
       CompanyName           := '';
       FileDescription       := '';
       FileVersion           := '';
       InternalName          := '';
       LegalCopyright        := '';
       LegalTradeMarks       := '';
       OriginalFilename      := '';
       ProductName           := '';
       ProductVersion        := '';
       Comments              := '';
       Language              := '';
       Translation           := '';
       FileVersionMajor      := 0;
       FileVersionMinor      := 0;
       FileVersionRelease    := 0;
       FileVersionBuild      := 0;
       ProductVersionMajor   := 0;
       ProductVersionMinor   := 0;
       ProductVersionRelease := 0;
       ProductVersionBuild   := 0;
    end;
    Size := GetFileVersionInfoSizeW(PWideChar(FileName), Handle);
    If Size > 0 then begin
       try    GetMem(Data, Size);
       except Exit;
       end;
       try
          If GetFileVersionInfow(PWideChar(FileName), Handle, Size, Data) then begin
             If VerQueryValue(Data, '\', Pointer(FixedFileInfo), Len) then begin
                VersionInfo.FileVersionMajor      := HiWord(FixedFileInfo^.dwFileVersionMS);
                VersionInfo.FileVersionMinor      := LoWord(FixedFileInfo^.dwFileVersionMS);
                VersionInfo.FileVersionRelease    := HiWord(FixedFileInfo^.dwFileVersionLS);
                VersionInfo.FileVersionBuild      := LoWord(FixedFileInfo^.dwFileVersionLS);
                VersionInfo.ProductVersionMajor   := HiWord(FixedFileInfo^.dwProductVersionMS);
                VersionInfo.ProductVersionMinor   := LoWord(FixedFileInfo^.dwProductVersionMS);
                VersionInfo.ProductVersionRelease := HiWord(FixedFileInfo^.dwProductVersionLS);
                VersionInfo.ProductVersionBuild   := LoWord(FixedFileInfo^.dwProductVersionLS);
                VersionInfo.FileVersion := IntToStr(HiWord(FixedFileInfo^.dwFileVersionMS)) + '.' + IntToStr(LoWord(FixedFileInfo^.dwFileVersionMS)) + '.' +
                                           IntToStr(HiWord(FixedFileInfo^.dwFileVersionLS)) + '.' + IntToStr(LoWord(FixedFileInfo^.dwFileVersionLS))
             end;

             If VerQueryValueW(Data, '\VarFileInfo\Translation', Buffer, Len) then begin
                Translation             := IntToHex(PDWORD(Buffer)^, 8);
                Translation             := Copy(Translation, 5, 4) + Copy(Translation, 1, 4);
                VersionInfo.Translation := '$' + Copy(Translation, 1, 4);
                SetLength(VersionInfo.Language, 64);
                SetLength(VersionInfo.Language, VerLanguageNameW(StrToIntDef('$' + Copy(Translation, 1, 4), $0409), PWideChar(VersionInfo.Language), 64));
             end;

             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\CompanyName'),      Buffer, Len) then VersionInfo.CompanyName      := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\FileDescription'),  Buffer, Len) then VersionInfo.FileDescription  := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\FileVersion'),      Buffer, Len) then VersionInfo.FileVersion      := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\InternalName'),     Buffer, Len) then VersionInfo.InternalName     := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\LegalCopyright'),   Buffer, Len) then VersionInfo.LegalCopyright   := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\LegalTradeMarks'),  Buffer, Len) then VersionInfo.LegalTradeMarks  := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\OriginalFilename'), Buffer, Len) then VersionInfo.OriginalFilename := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\ProductName'),      Buffer, Len) then VersionInfo.ProductName      := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\ProductVersion'),   Buffer, Len) then VersionInfo.ProductVersion   := PWideChar(Buffer);
             If VerQueryValueW(Data, PWideChar('\StringFileInfo\' + Translation + '\Comments'),         Buffer, Len) then VersionInfo.Comments         := PWideChar(Buffer);
             Result := True;
          end;
       finally
          FreeMem(Data);
       end;
    end;
end;


end.

