program Viewer;

uses
  Vcl.Forms,
  View in 'View.pas' {V};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TV, V);
  Application.Run;
end.
