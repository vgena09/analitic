unit FunVcl;

interface

uses
   Winapi.Windows, Winapi.Messages,
   System.SysUtils, System.Variants, System.Classes,
   Vcl.Forms, Vcl.Controls, Vcl.Dialogs, Vcl.ExtCtrls;

   function  LoadSubForm(ControlForm: TPanel; SubForm: TFormClass; const Update: Boolean): TForm;
   procedure ErrMsg(const Str: String);
   function  ForegroundWindow: Boolean;
   procedure BeginScreenUpdate(Hwnd: THandle);
   procedure EndScreenUpdate(Hwnd: THandle; const Erase: Boolean);


implementation

uses FunInfo;


{==============================================================================}
{============================= «¿√–”« ¿ —”¡‘Œ–Ã€ ==============================}
{==============================================================================}
function LoadSubForm(ControlForm: TPanel; SubForm: TFormClass; const Update: Boolean): TForm;
begin
    {»ÌËˆË‡ÎËÁ‡ˆËˇ}
    Result:=nil;
    If Update then LockWindowUpdate(ControlForm.Handle);
    try
       if (ControlForm.ControlCount = 0) or not (ControlForm.Controls[0] is SubForm) then begin
          if ControlForm.ControlCount > 0 then begin
             TForm(ControlForm.Controls[0]).Hide;
             TForm(ControlForm.Controls[0]).Close;
             ControlForm.Controls[0].Free;
          end;
          if Assigned(SubForm) then begin
             Result := SubForm.Create(nil);
             Result.Hide;
             Result.BorderStyle := bsNone;
             Result.Parent := ControlForm;
             Result.Align := alClient;
             Result.Show;
          end;
       end;
    finally If Update then LockWindowUpdate(0);
    end;
end;


{==============================================================================}
{=========================   —ŒŒ¡Ÿ≈Õ»≈ Œ¡ Œÿ»¡ ≈   ============================}
{==============================================================================}
procedure ErrMsg(const Str: String);
begin
    ForegroundWindow;
    MessageDlg(STR+Chr(13)+Chr(10)+Chr(13)+Chr(10)+
               ' ÓÌÚ‡ÍÚÌ‡ˇ ËÌÙÓÏ‡ˆËˇ:'+Chr(13)+Chr(10)+
               GetProgLegalCopyright+Chr(13)+Chr(10)+
               GetProgComments, mtError, [mbOk], 0);
end;


{==============================================================================}
{===================   Œ ÕŒ œ–Œ√–¿ÃÃ€ Õ¿ œ≈–≈ƒÕ»… œÀ¿Õ    =====================}
{==============================================================================}
function ForegroundWindow: Boolean;
begin
    Keybd_event(0,0,0,0);
    Result:=SetForegroundWindow(GetLastActivePopup(Application.Handle));
end;


{==============================================================================}
{==============  ” — “ – ¿ Õ ﬂ ≈ “   Ã ≈ – ÷ ¿ Õ » ≈   Œ   Õ ¿  ===============}
{==============================================================================}
procedure BeginScreenUpdate(Hwnd: THandle);
begin
   if (hwnd = 0) then hwnd := Application.MainForm.Handle;
   SendMessage(hwnd, WM_SETREDRAW, 0, 0);
end;

procedure EndScreenUpdate(Hwnd: THandle; const Erase: Boolean);
begin
    If (hwnd = 0) then hwnd := Application.MainForm.Handle;
    SendMessage(hwnd, WM_SETREDRAW, 1, 0);
    RedrawWindow(hwnd, nil, 0, RDW_FRAME + RDW_INVALIDATE + RDW_ALLCHILDREN + RDW_NOINTERNALPAINT);
    If (erase) then InvalidateRect(hwnd, nil, True);
end;

end.

