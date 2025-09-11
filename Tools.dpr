program Tools;

uses
  Vcl.Forms,
  Tool in 'Tool.pas' {T},
  Func in 'Func.pas',
  Tool_Mail in 'Tool_Mail.pas' {TM},
  Tool_Reception in 'Tool_Reception.pas' {TR},
  Tool_Company in 'Tool_Company.pas' {TC};

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TT, T);
  Application.Run;
end.
