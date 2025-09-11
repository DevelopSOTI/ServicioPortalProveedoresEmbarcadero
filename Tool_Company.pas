unit Tool_Company;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Menus, Data.Win.ADODB, Data.DB,
  Vcl.StdCtrls, Vcl.Grids, Vcl.DBGrids;

type
  TTC = class(TForm)
    Label4: TLabel;
    Label2: TLabel;
    DBGridEmpresas: TDBGrid;
    Conexion_MySQL: TADOConnection;
    Select_Empresas: TADOQuery;
    DataSource_Empresas: TDataSource;
    Command_Update: TADOCommand;
    Select_Diferencia: TADOQuery;
    PopupMenu1: TPopupMenu;
    Autorizar1: TMenuItem;
    Bloquear1: TMenuItem;
    Permitediferencias1: TMenuItem;
    Rechazadiferencias1: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure RESTART_SRV_BTNClick(Sender: TObject);
    procedure Autorizar1Click(Sender: TObject);
    procedure Bloquear1Click(Sender: TObject);
    procedure Permitediferencias1Click(Sender: TObject);
    procedure Rechazadiferencias1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  TC: TTC;

  empresas_autorizadas :Integer;
  empresas_limite :Integer;

const
  CAPTION_MESSAGE :PChar = 'Mensaje del modulo de empresas';

implementation

uses
  Tool, Func;

{$R *.dfm}

procedure TTC.FormCreate(Sender: TObject);
begin
  empresas_autorizadas := 0;
  empresas_limite := 1;
end;

procedure TTC.FormActivate(Sender: TObject);
  var
    ConnectionString :string;
begin
  ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;';
  ConnectionString := ConnectionString + 'UID=' + T.V_MYSQL_USER + ';';
  ConnectionString := ConnectionString + 'PORT=' + T.V_MYSQL_PORT + ';';
  ConnectionString := ConnectionString + 'DATABASE=' + T.V_MYSQL_DATA + ';';
  ConnectionString := ConnectionString + 'SERVER=' + T.V_MYSQL_SERV + ';';
  ConnectionString := ConnectionString + 'PASSWORD=' + T.V_MYSQL_PASS + ';';



  try
    Conexion_MySQL.ConnectionString := ConnectionString;
    Conexion_MySQL.Connected := True;
  except
    MessageBox( 0, 'No fue posible establecer conexión con el portal.', CAPTION_MESSAGE, MB_ICONERROR );
    Exit;
  end;



  try
    Select_Empresas.Active := True;
    Select_Empresas.First;
    while not Select_Empresas.Eof do
      begin
        if ( Select_Empresas.FieldByName('EMP_ESTATUS').AsString = 'Autorizada' ) then
          begin
            Inc( empresas_autorizadas );
          end;

        Select_Empresas.Next;
      end;
    Select_Empresas.First;

    { Select_Diferencia.Active := False;
    Select_Diferencia.SQL.Clear;
    Select_Diferencia.SQL.Add( 'SELECT PARAM_VALOR FROM PARAMETROS WHERE PARAM_CLAVE = ''TOLERANCIA''' );
    Select_Diferencia.Active := True;
    Select_Diferencia.First;
    TOLERANCIA_TXT.Text := Select_Diferencia.FieldByName('PARAM_VALOR').AsString; }
  except
    MessageBox( 0, 'No fue posible obtener la lista de empresas ni los parametros de tolerancia del portal.', CAPTION_MESSAGE, MB_ICONERROR );
  end;



  // Conexion_MySQL.Connected := False;
end;





procedure TTC.Autorizar1Click(Sender: TObject);
  var
    i :Integer;
    update :String;
begin
  if ( DBGridEmpresas.SelectedRows.Count > empresas_limite ) then
    begin
      MessageBox( 0, PChar( 'Solo se permiten dar de alta ' + IntToStr( empresas_limite ) + ' empresas.' ), CAPTION_MESSAGE, MB_ICONWARNING );
      Exit;
    end;

  if ( ( DBGridEmpresas.SelectedRows.Count + empresas_autorizadas ) > empresas_limite ) then
    begin
      MessageBox( 0, PChar( 'Solo se permiten dar de alta ' + IntToStr( empresas_limite ) + ' empresas.' ), CAPTION_MESSAGE, MB_ICONWARNING );
      Exit;
    end;

  if DBGridEmpresas.SelectedRows.Count > 0 then
    begin
      with DBGridEmpresas.DataSource.DataSet do
        begin
          for i := 0 to DBGridEmpresas.SelectedRows.Count - 1 do
            begin
              GotoBookmark( Pointer( DBGridEmpresas.SelectedRows.Items[i] ) );
              update := 'UPDATE EMPRESAS_MSP SET EMP_ESTATUS = ''Autorizada'' WHERE EMP_ID = ' + Select_Empresas.FieldByName('EMP_ID').AsString;
              try
                Command_Update.CommandText := update;
                Command_Update.Execute;
              except
                on E:Exception do
                  begin
                    MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message + #13 + 'Hubo un error al intentar autorizar las empresas' ), 'Mensaje del modulo de empresas', MB_ICONERROR );
                  end;
              end;
            end;
        end;

      Select_Empresas.Active := False;
      Select_Empresas.Active := True;

      empresas_autorizadas := 0;

      Select_Empresas.First;
      while not Select_Empresas.Eof do
        begin
          if ( Select_Empresas.FieldByName('EMP_ESTATUS').AsString = 'Autorizada' ) then
            begin
              Inc( empresas_autorizadas );
            end;
          Select_Empresas.Next;
        end;
      Select_Empresas.First;
    end;
end;

procedure TTC.Bloquear1Click(Sender: TObject);
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
              update := 'UPDATE EMPRESAS_MSP SET EMP_ESTATUS = ''Bloqueada'' WHERE EMP_ID = ' + Select_Empresas.FieldByName('EMP_ID').AsString;
              try
                Command_Update.CommandText := update;
                Command_Update.Execute;
              except
                on E:Exception do
                  begin
                    MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message + #13 + 'Hubo un error al intentar bloquear las empresas' ), 'Mensaje del modulo de empresas', MB_ICONERROR );
                  end;
              end;
            end;
        end;
      Select_Empresas.Active := False;
      Select_Empresas.Active := True;

      empresas_autorizadas := 0;
      Select_Empresas.First;
      while not Select_Empresas.Eof do
        begin
          if ( Select_Empresas.FieldByName('EMP_ESTATUS').AsString = 'Autorizada' ) then
            begin
              Inc( empresas_autorizadas );
            end;
          Select_Empresas.Next;
        end;
      Select_Empresas.First;
    end;
end;

procedure TTC.Permitediferencias1Click(Sender: TObject);
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
              update := 'UPDATE EMPRESAS_MSP SET EMP_DIFERENCIA = ''S'' WHERE EMP_ID = ' + Select_Empresas.FieldByName('EMP_ID').AsString;
              try
                Command_Update.CommandText := update;
                Command_Update.Execute;
              except
                on E:Exception do
                  begin
                    MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message + #13 + 'Hubo un error al intentar permitir las diferencias' ), 'Mensaje del modulo de empresas', MB_ICONERROR );
                  end;
              end;
            end;
        end;
      Select_Empresas.Active := False;
      Select_Empresas.Active := True;
    end;
end;

procedure TTC.Rechazadiferencias1Click(Sender: TObject);
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

              update := 'UPDATE EMPRESAS_MSP SET EMP_DIFERENCIA = ''N'' ';
              update := update + 'WHERE EMP_ID = ' + Select_Empresas.FieldByName('EMP_ID').AsString;

              try
                Command_Update.CommandText := update;
                Command_Update.Execute;
              except
                on E:Exception do
                  begin
                    MessageBox( 0, PChar( '[' + E.ClassName + '] ' + E.Message + #13 + 'Hubo un error al intentar rechazar las diferencias' ), 'Mensaje del modulo de empresas', MB_ICONERROR );
                  end;
              end;
            end;
        end;

      Select_Empresas.Active := False;
      Select_Empresas.Active := True;
    end;
end;





procedure TTC.RESTART_SRV_BTNClick(Sender: TObject);
begin
  { if isInstalled( 'SyncService' ) then // CHECA SI EL SERVICIO ESTA INSTALADO
    begin
      StopService( 'SyncService' ); // DETIENE EL SERVICIO
      while not isRunning( 'SyncService' ) do // MIENTRAS NO ARRANQUE EL SERVICIO
        begin
          StartSrv( 'SyncService' ); // VA A INTENTAR ARRANCARLO
          MessageBox( 0, 'Servicio reiniciado', 'Mensaje de empresas', MB_ICONINFORMATION );
          RESTART_SRV_BTN.Enabled := False;
        end;
    end
  else
    begin
      MessageBox( 0, 'El servicio no esta instalado', 'Mensaje de empresas', MB_ICONERROR );
    end; // }
end;



end.
