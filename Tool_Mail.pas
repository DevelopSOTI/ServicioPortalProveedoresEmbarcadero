unit Tool_Mail;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.Win.ADODB, Data.DB, Vcl.StdCtrls;

type
  TTM = class(TForm)
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    EDIT_NOMBRE: TEdit;
    EDIT_ASUNTO: TEdit;
    EDIT_SERVIDOR: TEdit;
    EDIT_PUERTO: TEdit;
    EDIT_CORREO: TEdit;
    EDIT_PASSWORD: TEdit;
    Button1: TButton;
    Conexion_MySQL: TADOConnection;
    Select_Correo: TADOQuery;
    Command_Update: TADOCommand;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  TM: TTM;

implementation

uses
  Tool;

{$R *.dfm}

procedure TTM.Button1Click(Sender: TObject);
  var
    update :String;
begin
  update := 'UPDATE MAIL SET ';
  update := update + 'MAIL_NAME = ' + QuotedStr( EDIT_NOMBRE.Text ) + ', ';
  update := update + 'MAIL_TITLE = ' + QuotedStr( EDIT_ASUNTO.Text ) + ', ';
  update := update + 'MAIL_SMTP = ' + QuotedStr( EDIT_SERVIDOR.Text ) + ', ';
  update := update + 'MAIL_PORT = ' + QuotedStr( EDIT_PUERTO.Text ) + ', ';
  update := update + 'MAIL_FROM = ' + QuotedStr( EDIT_CORREO.Text ) + ', ';
  update := update + 'MAIL_PASS = ' + QuotedStr( EDIT_PASSWORD.Text );
  try
    Command_Update.CommandText := update;
    Command_Update.Execute;
    MessageBox( 0, '¡Configuración guardada!', 'Mensaje del modulo de correos', MB_ICONINFORMATION );
  except
    on E:Exception do
      begin
        MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message + #13 + 'Hubo un error al intentar actualizar la información' ), 'Mensaje del modulo de correos', MB_ICONERROR );
      end;
  end;
end;

procedure TTM.FormCreate(Sender: TObject);
  var
    mysql_serv, mysql_user, mysql_pass, mysql_data, mysql_port :String;
begin
  mysql_serv := T.MYSQL_SERV.Text;
  mysql_user := T.MYSQL_USER.Text;
  mysql_pass := T.MYSQL_PASS.Text;
  mysql_data := T.MYSQL_DATA.Text;
  mysql_port := T.MYSQL_PORT.Text;

  Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + mysql_user + ';PORT=' + mysql_port + ';DATABASE=' + mysql_data + ';SERVER=' + mysql_serv + ';PASSWORD=' + mysql_pass + ';';
  try
    Conexion_MySQL.Connected := True;
    Select_Correo.Active := True;
    Select_Correo.First;

    EDIT_NOMBRE.Text := Select_Correo.FieldByName('MAIL_NAME').AsString;
    EDIT_ASUNTO.Text := Select_Correo.FieldByName('MAIL_TITLE').AsString;
    EDIT_SERVIDOR.Text := Select_Correo.FieldByName('MAIL_SMTP').AsString;
    EDIT_PUERTO.Text := Select_Correo.FieldByName('MAIL_PORT').AsString;
    EDIT_CORREO.Text := Select_Correo.FieldByName('MAIL_FROM').AsString;
    EDIT_PASSWORD.Text := Select_Correo.FieldByName('MAIL_PASS').AsString;
  except
    on E : Exception do
      begin
        MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message ), 'Mensaje del modulo de correos', MB_ICONERROR );
        Exit;
      end;
  end;
end;

end.
