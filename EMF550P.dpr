program EMF550P;

uses
  Vcl.Forms,
  EMF550P.Main in 'EMF550P.Main.pas' {FormMain};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TFormMain, FormMain);
  Application.Run;
end.
