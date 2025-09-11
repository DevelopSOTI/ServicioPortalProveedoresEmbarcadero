unit Tool_Reception;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Data.Win.ADODB, Data.DB,
  Vcl.Grids, Vcl.DBGrids;

type
  TTR = class(TForm)
    DBGridEmpresas: TDBGrid;
    Conexion_MySQL: TADOConnection;
    Select_Dias: TADOQuery;
    DataSource_Empresas: TDataSource;
    Command_Update: TADOCommand;
    PopupMenu1: TPopupMenu;
    Autorizar1: TMenuItem;
    Bloquear1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure DBGridEmpresasDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Bloquear1Click(Sender: TObject);
    procedure Autorizar1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  TR: TTR;

implementation

uses
  Tool;

{$R *.dfm}

procedure TTR.Autorizar1Click(Sender: TObject);
  var
    i :Integer;
    update :String;
begin
  if DBGridEmpresas.SelectedRows.Count > 0 then
    begin
      with DBGridEmpresas.DataSource.DataSet do
        begin
          for i := 0 to DBGridEmpresas.SelectedRows.Count - 1 do
            begin
              GotoBookmark( Pointer( DBGridEmpresas.SelectedRows.Items[i] ) );
              update := 'UPDATE DIAS SET DIA_RECIBE = TRUE WHERE DIA_NUMERO = ' + Select_Dias.FieldByName('DIA_NUMERO').AsString;
              try
                Command_Update.CommandText := update;
                Command_Update.Execute;
              except
                on E:Exception do
                  begin
                    MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message + #13 + 'Hubo un error al intentar activar los dias' ), 'Mensaje del modulo de recepción', MB_ICONERROR );
                  end;
              end;
            end;
        end;
      Select_Dias.Active := False;
      Select_Dias.Active := True;
    end;
end;

procedure TTR.Bloquear1Click(Sender: TObject);
  var
    i :Integer;
    update :String;
begin
  if DBGridEmpresas.SelectedRows.Count > 0 then
    begin
      with DBGridEmpresas.DataSource.DataSet do
        begin
          for i := 0 to DBGridEmpresas.SelectedRows.Count - 1 do
            begin
              GotoBookmark( Pointer( DBGridEmpresas.SelectedRows.Items[i] ) );
              update := 'UPDATE DIAS SET DIA_RECIBE = FALSE WHERE DIA_NUMERO = ' + Select_Dias.FieldByName('DIA_NUMERO').AsString;
              try
                Command_Update.CommandText := update;
                Command_Update.Execute;
              except
                on E:Exception do
                  begin
                    MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message + #13 + 'Hubo un error al intentar bloquear los dias' ), 'Mensaje del modulo de recepción', MB_ICONERROR );
                  end;
              end;
            end;
        end;
      Select_Dias.Active := False;
      Select_Dias.Active := True;
    end;
end;

procedure TTR.DBGridEmpresasDrawColumnCell(Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
  var
    txt :String;
begin
  if ( Column.FieldName = 'DIA_RECIBE' ) then // SI ES EL CAMPO DEL ID DEL CLIENTE.
    begin
      txt := Select_Dias.FieldByName('DIA_RECIBE').AsString;

      if ( txt = '0' ) then
        begin
          txt := 'No';
        end;

      if ( txt = '1' ) then
        begin
          txt := 'Si';
        end;

      DBGridEmpresas.Canvas.TextRect( Rect, Rect.Left + 5, Rect.Top + 2, txt );
    end;
end;

procedure TTR.FormCreate(Sender: TObject);
  var
    mysql_serv,
    mysql_user,
    mysql_pass,
    mysql_data,
    mysql_port :String;
begin
  mysql_serv := T.MYSQL_SERV.Text;
  mysql_user := T.MYSQL_USER.Text;
  mysql_pass := T.MYSQL_PASS.Text;
  mysql_data := T.MYSQL_DATA.Text;
  mysql_port := T.MYSQL_PORT.Text;

  Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + mysql_user + ';PORT=' + mysql_port + ';DATABASE=' + mysql_data + ';SERVER=' + mysql_serv + ';PASSWORD=' + mysql_pass + ';';
  try
    Conexion_MySQL.Connected := True;
    Select_Dias.Active := True;
  except
    on E : Exception do
      begin
        MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message ), 'Mensaje del modulo de recepción', MB_ICONERROR );
        Exit;
      end;
  end;
end;

end.
