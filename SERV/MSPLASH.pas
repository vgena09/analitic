unit MSPLASH;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,
  Vcl.ExtCtrls, Vcl.Imaging.jpeg;

type
  TFSPLASH = class(TForm)
    Image1: TImage;
    LVersion: TLabel;
    LInfo: TLabel;
    LUser: TLabel;
    LBank: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public
    procedure Msg(const SText: String);
  end;

var
  FSPLASH : TFSPLASH;

implementation

uses FunConst, FunSys, FunInfo;

{$R *.DFM}


procedure TFSPLASH.FormCreate(Sender: TObject);
begin
    LUser.Caption    := LUser.Caption   +'  '+GetWinUser+' ';
    LVersion.Caption := LVersion.Caption+'  '+GetProgFileVersion +' ';
    LInfo.Caption    := '';
end;

procedure TFSPLASH.Msg(const SText: String);
begin
    LInfo.Caption := 'Загрузка:  '+SText+'  ';
    LInfo.Repaint;
end;


end.
