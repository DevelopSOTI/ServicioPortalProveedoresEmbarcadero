program Service;

uses
  SvCom_NTService,
  Main in 'Main.pas' {M: TNtService},
  Data in 'Data.pas' {D: TDataModule},
  Form in 'Form.pas' {F},
  Func in 'Func.pas',
  Func_Calcula in 'Func_Calcula.pas',
  Func_Catalogos in 'Func_Catalogos.pas',
  Func_Recepciones in 'Func_Recepciones.pas',
  Func_Creditos in 'Func_Creditos.pas',
  Func_Facturas_3_2 in 'Func_Facturas_3_2.pas',
  Func_Facturas_3_3 in 'Func_Facturas_3_3.pas';

{$R *.RES}

begin
  Application.Initialize;
  Application.CreateForm(TD, D);
  // Application.CreateForm(TM, M); // ACTIVAR PARA INSTALAR EL SERVICIO
  Application.Run;
end.
