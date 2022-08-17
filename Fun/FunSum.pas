unit FunSum;

interface
uses
   System.Classes, System.SysUtils, System.IniFiles, System.Variants,
   Vcl.Dialogs, Vcl.Controls,
   Data.DB, Data.Win.ADODB,
   IdGlobal,
   FunType;

function  SumTable(const STab, SYear, SMonth, SRegion: String): Boolean;

implementation

uses FunConst, MAIN,
     FunBD, FunText, FunDay, FunIO, FunSys, FunIni;

{==============================================================================}
{===================     СУММИРОВАНИЕ ДАННЫХ ТАБЛИЦЫ    =======================}
{==============================================================================}
{============   SRegion - название региона с измененной таблицей  =============}
{==============================================================================}
function SumTable(const STab, SYear, SMonth, SRegion: String): Boolean;
var FFMAIN        : TFMAIN;
    LRegion_      : TLRegion;
    Sr            : TSearchRec;
    DatPeriod     : TDate;
    IYear, IMonth : Word;
    IDRegion      : Integer;

    SListParent   : TStringList;
    IDParent      : Integer;
    SParentRegion : String;
    SPathDat      : String;
    IndParent     : Integer;

    ExistsSection : Boolean;
    S, S1, S2     : String;
    I             : Integer;

    function SumStep(const SRegionStep: String): Boolean;
    var FDat        : TMemIniFile;
        SList       : TStringList;
        SDat        : String;
        S, S0, SKey : String;
        I           : Integer;
        IVal        : Extended;
    begin
        {Инициализация}
        Result:=false;

        {Инициализируем dat-файл текущего региона}
        SDat:=FFMAIN.PathDat(SYear, SMonth, SRegionStep, STab);
        If Not FileExists(SDat) then Exit;
        {Секция должна существовать}
        If Not FFMAIN.IsExistsTab(SDat, STab) then Exit;

        FDat  := TMemIniFile.Create(SDat);
        SList := TStringList.Create;
        try
           {Запоминаем содержимое файла}
           FDat.ReadSectionValues(STab, SList);

           {Производим суммирование}
           For I:=0 to SList.Count-1 do begin
              S    := SList[I];                        //  S = '1 2=5'
              SKey := TokChar(S, '=');                 //  S = '5'
              {Если значение допустимо}
              If IsIntegerStr(S) then begin
                 {Ищем соответствующую родительскую запись}
                 IndParent:=SListParent.IndexOfName(SKey);
                 {Если родительская запись существует}
                 If IndParent>=0 then begin
                    S0:=SListParent.Values[SKey];
                    If IsIntegerStr(S0) then IVal:=StrToFloat(S0) else IVal:=0;
                    SListParent.Values[SKey]:=FloatToStr(StrToFloat(S)+IVal);
                 {Если родительская запись отсутствует}
                 end else begin
                    SListParent.Add(SKey+'='+S);
                 end;
              end;
           end;

           {Возвращаемый результат}
           Result:=true;
        finally
           SList.Free;
           FDat.Free;
        end;
    end;

begin
    {Инициализация}
    Result := false;
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));

    If Not IsIntegerStr(SYear) then Exit;
    IYear  :=StrToInt(SYear);
    IMonth := MonthStrToInd(SMonth);
    If IMonth<=0 then Exit;
    DatPeriod := IDatePeriod(IYear, IMonth);
    If DatPeriod=0 then Exit;

    S2 := CutModulStr (SRegion,     SUB_RGN1, SUB_RGN2);
    S1 := ReplModulStr(SRegion, '', SUB_RGN1, SUB_RGN2);

    SListParent := TStringList.Create;
    try
        {Индекс исходного узла}
        If Not FFMAIN.FindRegions(@LRegion_, true, -1, S1, -1, -1, DatPeriod) then Exit;
        IDRegion := LRegion_[0].FCounter;


        {**********************************************************************}
        {*****   Суммирование возможных свободных регионов   ******************}
        {**********************************************************************}
        If S2<>'' then begin
           {Название и индекс родительского узла}
           SParentRegion := S1;
           IDParent      := IDRegion;

           {Информатор}
           FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Расчет региона: '+STab+'! [ '+SParentRegion+' ]';

           {Путь к итоговому dat-файлу}
           SPathDat := FFMAIN.PathDat(SYear, SMonth, S1, STab);

           {Ищем первый свободный регион}
           S:=ExtractFilePath(SPathDat);
           try
              If FindFirst(S+S1+SUB_RGN1+'*', faDirectory, Sr) = 0 then begin
                 {Удаляем предыдущую итоговую секцию dat-файла}
                 FFMAIN.DelTab(SPathDat, STab);

                 {Производим расчет}
                 repeat S:=ExtractFileNameWithoutExt(Sr.Name);
                        SumStep(S);
                 until  FindNext(Sr) <> 0;

                 {Сохраняем обновленную секцию dat-файла родительского узла}
                 If Not FFMAIN.SetTabVal(SPathDat, STab, @SListParent, true) then Exit;
              end;
           finally FindClose(Sr); end;


        {**********************************************************************}
        {*****   Суммирование регионов   **************************************}
        {**********************************************************************}
        end else begin
           {Нестандартную таблицу пропускаем без ошибки}
           If FFMAIN.IsTableNoStandart(STab) then begin
              Result:=true;
              Exit;
           end;

           {Индекс родительского узла}
           IDParent:=FFMAIN.RegionsCounterToParent(IDRegion);
           If IDParent=-1 then Exit;

           {Если регион всего один, то нечего суммировать}
           If IDParent=0 then begin
              Result:=true;
              Exit;
           end;

           {Название родительского узла}
           SParentRegion:=FFMAIN.RegionsCounterToCaption(IDParent);
           If SParentRegion='' then Exit;

           {Информатор}
           FFMAIN.StatusBar.Panels[STATUS_MAIN].Text:=' Расчет региона: '+STab+'! [ '+SParentRegion+' ]';

           {Удаляем предыдущую итоговую секцию dat-файла}
           SPathDat:=FFMAIN.PathDat(SYear, SMonth, SParentRegion, STab);
           FFMAIN.DelTab(SPathDat, STab);

           {Выбираем регионы с одинаковым родительским узлом}
           If Not FFMAIN.FindRegions(@LRegion_, false, -1, '', IDParent, -1, DatPeriod) then Exit;

           {Поочередно суммируем выбранные регионы}
           ExistsSection := true;
           For I:=Low(LRegion_) to High(LRegion_) do begin
              {Если таблица региона отсутствует, то итоговый файл создать нельзя}
              ExistsSection := SumStep(LRegion_[I].FCaption);
              {Если принудительное суммирование, то ошибки пропускаем}
              If FFMAIN.IsTableNoStandart(STab) then ExistsSection:=true;
              If Not ExistsSection then Break;
           end;

           {Сохраняем обновленную секцию dat-файла родительского узла}
           If ExistsSection then begin
              If Not FFMAIN.SetTabVal(SPathDat, STab, @SListParent, true) then Exit;
           {Удаляем секцию dat-файла родительского узла}
           end else begin
              If Not FFMAIN.DelTab(SPathDat, STab) then Exit;
           end;
        end;
        {**********************************************************************}
     finally
        SListParent.Free;
        SetLength(LRegion_, 0);
     end;

     {Рекурсивный вызов расчета следующего уровня}
     If IDParent <> IMAIN_REGION then Result := SumTable(STab, SYear, SMonth, SParentRegion)
                                 else Result := true;
end;

end.

