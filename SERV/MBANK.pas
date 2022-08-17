unit MBANK;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  MAIN, Vcl.ActnList, Vcl.ExtCtrls, Vcl.StdCtrls, Vcl.Buttons;

type
  TFBANK = class(TForm)
    PBottom: TPanel;
    ActionList1: TActionList;
    AOk: TAction;
    ACancel: TAction;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    LBox: TListBox;
    procedure FormCreate(Sender: TObject);
    procedure AOkExecute(Sender: TObject);
    procedure ACancelExecute(Sender: TObject);
  private
    FFMAIN   : TFMAIN;
  public
    function Execute: Integer;
  end;

var
  FBANK: TFBANK;

implementation

uses FunConst, FunSys;

{$R *.dfm}

{==============================================================================}
{===========================   ÑÎÇÄÀÍÈÅ ÔÎÐÌÛ   ===============================}
{==============================================================================}
procedure TFBANK.FormCreate(Sender: TObject);
var I: Integer;
begin
    {Èíèöèàëèçàöèÿ}
    FFMAIN := TFMAIN(GlFindComponent('FMAIN'));
    {Ñïèñîê áàíêîâ äàííûõ}
    LBox.Clear;
    LBox.Items.Add('Àâòîâûáîð');
    For I := Low(FFMAIN.LBank) to High(FFMAIN.LBank) do LBox.Items.Add('Áàíê äàííûõ: '+IntToStr(FFMAIN.LBank[I].ID)); 
    LBox.ItemIndex := 0;
end;


{==============================================================================}
{============================   ÂÍÅØÍßß ÔÓÍÊÖÈß  ==============================}
{==============================================================================}
function TFBANK.Execute: Integer;
begin
    If Length(FFMAIN.LBank) > 1 then begin
       If ShowModal = mrOk then begin
          Result := LBox.ItemIndex - 1;
       end else begin
          Result := -2;
       end;
    end else begin
       Result := -1;
    end;
end;


{==============================================================================}
{=============================    ACTION: OK    ===============================}
{==============================================================================}
procedure TFBANK.AOkExecute(Sender: TObject);
begin ModalResult := mrOk; end;


{==============================================================================}
{=============================  ACTION: CANCEL  ===============================}
{==============================================================================}
procedure TFBANK.ACancelExecute(Sender: TObject);
begin ModalResult := mrCancel; end;


end.
