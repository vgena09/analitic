program Analitic;

uses
  Forms,
  MAIN in 'MAIN.pas' {FMAIN},
  FunBD in 'Fun\FunBD.pas',
  FunText in 'Fun\FunText.pas',
  FunType in 'Fun\FunType.pas',
  FunIO in 'Fun\FunIO.pas',
  FunIni in 'Fun\FunIni.pas',
  FunConst in 'Fun\FunConst.pas',
  FunVerify in 'Fun\FunVerify.pas',
  FunSum in 'Fun\FunSum.pas',
  FunVcl in 'Fun\COMMON\FunVcl.pas',
  FunInfo in 'Fun\COMMON\FunInfo.pas',
  FunFiles in 'Fun\COMMON\FunFiles.pas',
  FunDay in 'Fun\COMMON\FunDay.pas',
  FunSys in 'Fun\FunSys.pas',
  FunExcel in 'Fun\COMMON\FunExcel.pas',
  FunTree in 'Fun\COMMON\FunTree.pas',
  MSPLASH in 'SERV\MSPLASH.pas' {FSPLASH},
  MTITLE in 'SERV\MTITLE.pas' {FTITLE},
  MEXPORT in 'SERV\MEXPORT.pas' {FEXPORT},
  MIMPORT_XLS in 'SERV\MIMPORT_XLS.pas',
  MIMPORT_RAR in 'SERV\MIMPORT_RAR.pas' {FIMPORT_RAR},
  MBASE in 'SERV\MBASE.pas',
  MSET in 'SERV\MSET.pas' {FSET},
  MREPORT in 'SERV\MREPORT.pas',
  MBLANK in 'SERV\MBLANK.pas' {FBLANK},
  MPARAM in 'SERV\MPARAM.pas' {FPARAM},
  MINFO in 'SERV\MINFO.pas' {FINFO},
  MBDREGION in 'SERV\MBDREGION.pas' {FBDREGION},
  MBANK in 'SERV\MBANK.pas' {FBANK},
  MKOD in 'SERV\KOD\MKOD.pas' {FKOD},
  MKOD_FORMULA in 'SERV\KOD\MKOD_FORMULA.pas' {FKOD_FORMULA},
  MKOD_CORRECT in 'SERV\KOD\MKOD_CORRECT.pas' {FKOD_CORRECT},
  MMATRIX in 'SERV\MATRIX\MMATRIX.pas' {FMATRIX},
  MABOUT in 'SERV\MABOUT.pas' {FABOUT},
  MCLOSE in 'SERV\MCLOSE.pas' {FCLOSE};

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '—ËÒÚÂÏ‡ "¿Õ¿À»“» ¿"';
  Application.CreateForm(TFMAIN, FMAIN);
  Application.Run;
end.
