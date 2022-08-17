unit MABOUT;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.StdCtrls;

type
  TFABOUT = class(TForm)
    BtnOk: TButton;
    Img: TImage;
    LCaption: TLabel;
    LVersion: TLabel;
    LAuthor: TLabel;
    LContact: TLabel;
    Bevel4: TBevel;
    Label1: TLabel;
    procedure FormCreate(Sender: TObject);
  private

  public

  end;

var
  FABOUT: TFABOUT;

implementation

uses FunConst, FunSys, FunInfo;

{$R *.dfm}

procedure TFABOUT.FormCreate(Sender: TObject);
begin
    LVersion.Caption := 'Версия программы: '+GetProgProductVersion;
    LContact.Caption := MSGQWEST;
end;

end.
