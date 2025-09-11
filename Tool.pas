unit Tool;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Samples.Spin, Winapi.ShellAPI, System.Win.Registry,
  WinSvc, Vcl.ExtCtrls, IBX.IBDatabase, ADODB, Vcl.Menus;

type
  TT = class(TForm)
    MainMenu: TMainMenu;
    Herramientas1: TMenuItem;
    Empresas1: TMenuItem;
    Correo1: TMenuItem;
    Diasderecepcin1: TMenuItem;
    GroupBox1: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    MICRO_SERV: TEdit;
    MICRO_ROOT: TEdit;
    MICRO_USER: TEdit;
    MICRO_PASS: TEdit;
    GroupBox2: TGroupBox;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label14: TLabel;
    Label5: TLabel;
    MYSQL_SERV: TEdit;
    MYSQL_DATA: TEdit;
    MYSQL_USER: TEdit;
    MYSQL_PASS: TEdit;
    MYSQL_PORT: TEdit;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    Label13: TLabel;
    BTN_Install: TButton;
    BTN_Uninstall: TButton;
    BTN_Start: TButton;
    BTN_Stop: TButton;
    EDIT_TIME: TEdit;
    Panel1: TPanel;
    BTN_ACCEPT: TButton;
    BTN_CANCEL: TButton;
    lblTimer: TLabel;
    Label8: TLabel;
    checkAutomatico: TCheckBox;
    checkEnviaCorreo: TCheckBox;
    checkCierraConfig: TCheckBox;
    comboTime: TComboBox;
    procedure BTN_ACCEPTClick(Sender: TObject);
    procedure BTN_CANCELClick(Sender: TObject);
    procedure BTN_InstallClick(Sender: TObject);
    procedure BTN_UninstallClick(Sender: TObject);
    procedure BTN_StartClick(Sender: TObject);
    procedure BTN_StopClick(Sender: TObject);
    procedure Empresas1Click(Sender: TObject);
    procedure Correo1Click(Sender: TObject);
    procedure Diasderecepcin1Click(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }

    V_MICRO_SERV,
    V_MICRO_ROOT,
    V_MICRO_USER,
    V_MICRO_PASS,

    V_MYSQL_SERV,
    V_MYSQL_DATA,
    V_MYSQL_USER,
    V_MYSQL_PASS,
    V_MYSQL_PORT :string;
  end;

var
  T: TT;

  LOAD_ROOT :string;
  Troubles :Integer;

implementation

uses
  Func, Tool_Company, Tool_Mail, Tool_Reception;

{$R *.dfm}

{$REGION 'VALIDA_MICROSIP_CONNECT - VALIDA LA CONEXIÓN A MICROSIP'}
Function VALIDA_MICROSIP_CONNECT( serv, root, user, pass :string ):Boolean;
  var
    connect :TIBDatabase;
begin
  connect := TIBDatabase.Create( T );
  connect.LoginPrompt := False;
  connect.DatabaseName := serv + ':' + root + 'System\Config.FDB';
  connect.Params.Values['user_name'] := user;
  connect.Params.Values['password'] := pass;

  try
    connect.Connected := True;

    Result := True;
  except
    if ( MessageBox( 0, PChar( 'La configuración actual no permitira al replicador conectarse a Microsip.' + #13#13 + '¿Desea continuar?' ), 'Mensaje de configuración', MB_YESNO OR MB_ICONWARNING ) = ID_YES ) then
      begin
        Inc( Troubles );
        Result := True;
      end
    else
      begin
        Inc( Troubles );
        Result := False;
      end;
  end;

  connect.Connected := False;
  connect.Destroy;
end;
{$ENDREGION}

{$REGION 'VALIDA_MYSQL_CONNECT - VALIDA LA CONEXIÓN AL PORTAL DE PROVEEDORES'}
Function VALIDA_MYSQL_CONNECT( user, port, data, serv, pass, auto :string ):Boolean;
  var
    connect :TADOConnection;
    update :TADOCommand;
begin
  connect := TADOConnection.Create( Application );
  connect.LoginPrompt := False;
  connect.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + user + ';PORT=' + port + ';DATABASE=' + data + ';SERVER=' + serv + ';PASSWORD=' + pass + ';';

  try
    connect.Connected := True;

    update := TADOCommand.Create( connect );
    update.Connection := connect;
    update.CommandText := 'UPDATE PARAMETROS SET PARAM_VALOR = ''' + auto + ''' WHERE PARAM_CLAVE = ''APLICA_DIR''';
    update.Execute;

    Result := True;
  except
    if ( MessageBox( 0, PChar( 'La configuración actual no permitira al replicador conectarse al portal.' + #13#13 + '¿Desea continuar?' ), 'Mensaje de configuración', MB_YESNO OR MB_ICONWARNING ) = ID_YES ) then
      begin
        Inc( Troubles );
        Result := True;
      end
    else
      begin
        Inc( Troubles );
        Result := False;
      end;
  end;

  connect.Connected := False;
  connect.Destroy;
end;
{$ENDREGION}

{$REGION 'SAVE_CONFIG_INI - GUARDA EL TIEMPO DE REPLICACIÓN EN EL ARCHIVO Timer.ini'}
Function SAVE_CONFIG_INI( time :string ):Boolean;
  var
    Fichero :TStringList;
begin
  try
    Fichero := TStringList.Create;
    Fichero.Add( time );
    Fichero.SaveToFile( 'Timer.ini' );

    Result := True;
  except
    MessageBox( 0, PChar( 'No fue posible guardar los cambios porque el archivo ''Timer.ini'' esta siendo utilizado.' ), 'Mensaje de configuración', MB_ICONWARNING );

    Result := False;
  end;
end;
{$ENDREGION}

{$REGION 'SAVE_CONFIG - GUARDA LAS CONFIGURACIONES EN LOS REGISTROS DE WINDOWS'}
// Function SAVE_CONFIG( MICRO_SERV, MICRO_ROOT, MICRO_USER, MICRO_PASS, MYSQL_SERV, MYSQL_DATA, MYSQL_USER, MYSQL_PASS, MYSQL_PORT, CLOSE_AUTO, MAILS_SEND :string ):Boolean;
Function SAVE_CONFIG( MICRO_SERV, MICRO_ROOT, MICRO_USER, MICRO_PASS, MYSQL_SERV, MYSQL_DATA, MYSQL_USER, MYSQL_PASS, MYSQL_PORT, MAILS_SEND :string ):Boolean;
  var
    Reg :TRegistry;
begin
  Reg := TRegistry.Create( KEY_WRITE or KEY_WOW64_64KEY );
  Reg.RootKey := HKEY_LOCAL_MACHINE;
  if ( Reg.OpenKey( 'SOFTWARE\SOTI\Service Portal', True ) ) then
    begin
      Reg.WriteString( 'MICRO_SERV', MICRO_SERV );
      Reg.WriteString( 'MICRO_ROOT', MICRO_ROOT );
      Reg.WriteString( 'MICRO_USER', MICRO_USER );
      Reg.WriteString( 'MICRO_PASS', MICRO_PASS );

      Reg.WriteString( 'MYSQL_SERV', MYSQL_SERV );
      Reg.WriteString( 'MYSQL_DATA', MYSQL_DATA );
      Reg.WriteString( 'MYSQL_USER', MYSQL_USER );
      Reg.WriteString( 'MYSQL_PASS', MYSQL_PASS );
      Reg.WriteString( 'MYSQL_PORT', MYSQL_PORT );

      Reg.WriteString( 'MAILS_SEND', MAILS_SEND );
      // Reg.WriteString( 'CLOSE_AUTO', CLOSE_AUTO );

      Reg.WriteString( 'MODE_APPLI', 'S' );

      Reg.WriteString( 'MODE_TIMER', IntToStr( T.comboTime.ItemIndex ) );

      Result := True;
    end
  else
    begin
      Result := False;
    end;
  Reg.CloseKey;
  Reg.Free;
end;
{$ENDREGION}





procedure TT.FormCreate(Sender: TObject);
begin
  EDIT_TIME.Text := '';
  lblTimer.Caption := '';

  LOAD_ROOT := GetCurrentDir;
end;

procedure TT.FormActivate(Sender: TObject);
  var
    ConnectionString :string;

    MYSQL :TADOConnection;
    SELECT :TADOQuery;

    Reg :TRegistry;
    F :TextFile;
    Segundos :Integer;

    ExistReg :Boolean;
    ExistFile :Boolean;
begin
  if ( isRunning( 'SyncService' ) ) or ( isInstalled( 'SyncService' ) ) then
    begin
      BTN_Stop.Enabled := True;
      BTN_Uninstall.Enabled := True;
    end;

  {$REGION 'LEEMOS LOS REGISTROS GUARDADOS (SOFTWARE\SOTI\Service Portal)'}
  Reg := TRegistry.Create( KEY_READ or KEY_WOW64_64KEY );
  Reg.RootKey := HKEY_LOCAL_MACHINE;
  if ( Reg.KeyExists( 'SOFTWARE\SOTI\Service Portal' ) ) then
    begin
      if ( Reg.OpenKey( 'SOFTWARE\SOTI\Service Portal', False ) ) then
        begin
          // CARGAMOS LOS DATOS DE CONEXIÓN DE MICROSIP
          MICRO_SERV.Text := Reg.ReadString('MICRO_SERV');
          MICRO_ROOT.Text := Reg.ReadString('MICRO_ROOT');
          MICRO_USER.Text := Reg.ReadString('MICRO_USER');
          MICRO_PASS.Text := Reg.ReadString('MICRO_PASS'); // }

          V_MICRO_SERV := Reg.ReadString('MICRO_SERV');
          V_MICRO_ROOT := Reg.ReadString('MICRO_ROOT');
          V_MICRO_USER := Reg.ReadString('MICRO_USER');
          V_MICRO_PASS := Reg.ReadString('MICRO_PASS');

          // CARGAMOS LOS DATOS DE CONEXIÓN DEL PORTAL
          MYSQL_SERV.Text := Reg.ReadString('MYSQL_SERV');
          MYSQL_DATA.Text := Reg.ReadString('MYSQL_DATA');
          MYSQL_USER.Text := Reg.ReadString('MYSQL_USER');
          MYSQL_PASS.Text := Reg.ReadString('MYSQL_PASS');
          MYSQL_PORT.Text := Reg.ReadString('MYSQL_PORT'); // }

          V_MYSQL_SERV := Reg.ReadString('MYSQL_SERV');
          V_MYSQL_DATA := Reg.ReadString('MYSQL_DATA');
          V_MYSQL_USER := Reg.ReadString('MYSQL_USER');
          V_MYSQL_PASS := Reg.ReadString('MYSQL_PASS');
          V_MYSQL_PORT := Reg.ReadString('MYSQL_PORT');

          // REVISAMOS SI VA A ENVIAR O NO CORREOS EL SERVICIO
          if Reg.ReadString('MAILS_SEND') = 'True' then
            begin
              checkEnviaCorreo.Checked := True;
            end
          else
            begin
              checkEnviaCorreo.Checked := False;
            end;

          // REVISAMOS SI VA A CERRAR AUTOMATICAMENTE LA APLICACIÓN (DESACTIVADO)
          { if Reg.ReadString('CLOSE_AUTO') = 'True' then
            begin
              checkCierraConfig.Checked := True;
            end
          else
            begin
              checkCierraConfig.Checked := False;
            end; // }

          // INDICAMOS SI EL LAPSO DE TIEMPO ERA DE SEGUNDOS, MINUTOS O HORAS
          try
            comboTime.ItemIndex := StrToInt( Reg.ReadString('MODE_TIMER') );
          except
            comboTime.ItemIndex := 0;
          end;

          // INDICAMOS QUE SI SE ENCONTRARON LOS REGISTROS
          ExistReg := True;
        end
      else
        begin
          // DESHABILITAMOS LOS ACCESOS A OTROS MODULOS E INDICAMOS QUE NO SE ENCONTRARON LOS REGISTROS
          Empresas1.Enabled := False;
          Correo1.Enabled := False;
          Diasderecepcin1.Enabled := False;

          ExistReg := False;
        end;
    end
  else
    begin
      // DESHABILITAMOS LOS ACCESOS A OTROS MODULOS E INDICAMOS QUE NO SE ENCONTRARON LOS REGISTROS
      Empresas1.Enabled := False;
      Correo1.Enabled := False;
      Diasderecepcin1.Enabled := False;

      ExistReg := False;
    end;
  Reg.CloseKey;
  Reg.Free;
  {$ENDREGION}

  {$REGION 'REVISAMOS QUE EL ARCHIVO "Timer.ini" exista'}
  if ( FileExists( 'Timer.ini' ) ) then
    begin
      try
        AssignFile( F, 'Timer.ini' );
        Reset( F );
        Readln( F, Segundos );
        CloseFile( F );

        case comboTime.ItemIndex of
          1 : Segundos := Round( Segundos / 60 );
          2 : Segundos := Round( Segundos / 3600 );
        end;

        EDIT_TIME.Text := IntToStr( Segundos );
        ExistFile := True;
      except
        lblTimer.Caption := 'El archivo ''Timer.ini'' tiene valores invalidos.';
        EDIT_TIME.Text := IntToStr( 0 );
        ExistFile := False;
      end;
    end
  else
    begin
      lblTimer.Caption := 'El archivo ''Timer.ini'' no se encuentra en la carpeta del sistema.';
      EDIT_TIME.Text := IntToStr( 0 );
      ExistFile := False;
    end;
  {$ENDREGION}

  T.Update;



  if ExistReg and ExistFile then
    begin
      try
        ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;';
        ConnectionString := ConnectionString + 'UID=' + V_MYSQL_USER + ';';
        ConnectionString := ConnectionString + 'PORT=' + V_MYSQL_PORT + ';';
        ConnectionString := ConnectionString + 'DATABASE=' + V_MYSQL_DATA + ';';
        ConnectionString := ConnectionString + 'SERVER=' + V_MYSQL_SERV + ';';
        ConnectionString := ConnectionString + 'PASSWORD=' + V_MYSQL_PASS + ';';

        MYSQL := TADOConnection.Create( Self );
        MYSQL.LoginPrompt := False;
        MYSQL.ConnectionString := ConnectionString;
        MYSQL.Connected := True;
      except
        MessageBox(0, PChar( 'No fue posible establecer conexión con el portal revise su configuración.'), 'Mensaje de configuración', MB_ICONWARNING);

        Empresas1.Enabled := False;
        Correo1.Enabled := False;
        Diasderecepcin1.Enabled := False;

        Exit;
      end;



      try
        SELECT := TADOQuery.Create( MYSQL );
        SELECT.Connection := MYSQL;
        SELECT.SQL.Add( 'SELECT PARAM_VALOR FROM PARAMETROS WHERE PARAM_CLAVE = ''APLICA_DIR''' );
        SELECT.Active := True;
        SELECT.First;

        if ( UpperCase( SELECT.FieldByName('PARAM_VALOR').AsString ) = 'TRUE' ) then
          begin
            checkAutomatico.Checked := True;
          end
        else
          begin
            checkAutomatico.Checked := False;
          end;

        BTN_Install.Enabled := True;
        BTN_Uninstall.Enabled := True;
        BTN_Start.Enabled := True;
        BTN_Stop.Enabled := True;
      except
        MessageBox( 0, PChar( 'No fue posible obtener la configuración actual.' ), 'Mensaje de configuración', MB_ICONWARNING );

        Empresas1.Enabled := False;
        Correo1.Enabled := False;
        Diasderecepcin1.Enabled := False;
      end;



      MYSQL.Connected := False;
    end;
end;





procedure TT.BTN_ACCEPTClick(Sender: TObject);
  var
    TIME :Integer;

    AUTO,
    // CLOSE_AUTO,
    MAILS_SEND :string;
begin
  if isRunning( 'SyncService' ) then
    begin
      MessageBox( 0, 'Para poder hacer cambios en la configuración es necesario detener el replicador antes.', 'Mensaje de configuración', MB_ICONWARNING );
      Exit;
    end;

  {$REGION 'VALIDAMOS QUE EL TIEMPO INDICADO DE EJECUCIÓN SEA UN VALOR NUMERICO ENTERO VALIDO Y MAYOR A 0'}
  if ( EDIT_TIME.Text <> '' ) then
    begin
      try
        TIME := StrToInt( EDIT_TIME.Text );

        case comboTime.ItemIndex of
          1: TIME := TIME * 60;
          2: TIME := TIME * 3600;
        end;

        if TIME <= 0 then
          begin
            MessageBox( 0, PChar( 'El tiempo de ejecución debe ser mayor a 0.' ), 'Mensaje de configuración', MB_ICONWARNING );
            EDIT_TIME.SetFocus;
            Exit;
          end;
      except
        MessageBox( 0, PChar( 'El tiempo de ejecución debe ser un valor numerico entero ''' + EDIT_TIME.Text + ''' no es un valor valido.' ), 'Mensaje de configuración', MB_ICONWARNING );
        EDIT_TIME.SetFocus;
        Exit;
      end;
    end
  else
    begin
      MessageBox( 0, 'No indico el periodo de ejecución del replicador.', 'Mensaje de configuración', MB_ICONWARNING );
      EDIT_TIME.SetFocus;
      Exit;
    end;
  {$ENDREGION}


  if ( MessageBox( 0, PChar( 'Se guardara la configuración actual.' + #13#13 + '¿Desea continuar?' ), 'Mensaje de configuración', MB_YESNO OR MB_ICONQUESTION ) = ID_YES ) then
    begin
      BTN_ACCEPT.Caption := 'Validando';
      BTN_ACCEPT.Cursor := crHourGlass;
      BTN_ACCEPT.Enabled := False;

      Troubles := 0;



      // SE GENERARA DE FORMA AUTOMATICA SI O NO
      if checkAutomatico.Checked then
        begin
          AUTO := 'TRUE';
        end
      else
        begin
          AUTO := 'FALSE';
        end;

      // SE CERRARA DE FORMA AUTOMATICA SI O NO
      { if checkCierraConfig.Checked then
        begin
          CLOSE_AUTO := 'True';
        end
      else
        begin
          CLOSE_AUTO := 'False';
        end; // }

      // ENVIARA CORREO A LOS PROVEEDORES SI O NO
      if checkEnviaCorreo.Checked then
        begin
          MAILS_SEND := 'True';
        end
      else
        begin
          MAILS_SEND := 'False';
        end;



      if ( VALIDA_MICROSIP_CONNECT( MICRO_SERV.Text, MICRO_ROOT.Text, MICRO_USER.Text, MICRO_PASS.Text ) = True ) then
        begin
          if ( VALIDA_MYSQL_CONNECT( MYSQL_USER.Text, MYSQL_PORT.Text, MYSQL_DATA.Text, MYSQL_SERV.Text, MYSQL_PASS.Text, AUTO ) = True ) then
            begin
              if ( SAVE_CONFIG_INI( IntToStr( TIME ) ) = True ) then
                begin
                  // if ( SAVE_CONFIG( MICRO_SERV.Text, MICRO_ROOT.Text, MICRO_USER.Text, MICRO_PASS.Text, MYSQL_SERV.Text, MYSQL_DATA.Text, MYSQL_USER.Text, MYSQL_PASS.Text, MYSQL_PORT.Text, CLOSE_AUTO, MAILS_SEND ) = True ) then
                  if ( SAVE_CONFIG( MICRO_SERV.Text, MICRO_ROOT.Text, MICRO_USER.Text, MICRO_PASS.Text, MYSQL_SERV.Text, MYSQL_DATA.Text, MYSQL_USER.Text, MYSQL_PASS.Text, MYSQL_PORT.Text, MAILS_SEND ) = True ) then
                    begin
                      V_MICRO_SERV := MICRO_SERV.Text;
                      V_MICRO_ROOT := MICRO_ROOT.Text;
                      V_MICRO_USER := MICRO_ROOT.Text;
                      V_MICRO_PASS := MICRO_PASS.Text;

                      V_MYSQL_SERV := MYSQL_SERV.Text;
                      V_MYSQL_DATA := MYSQL_DATA.Text;
                      V_MYSQL_USER := MYSQL_USER.Text;
                      V_MYSQL_PASS := MYSQL_PASS.Text;
                      V_MYSQL_PORT := MYSQL_PORT.Text;

                      if Troubles = 0 then
                        begin
                          MessageBox( 0, 'Configuración guardada satisfactoriamente.', 'Mensaje de configuración', MB_ICONINFORMATION );

                          Empresas1.Enabled := True;
                          Correo1.Enabled := True;
                          Diasderecepcin1.Enabled := True;

                          BTN_Install.Enabled := True;
                          BTN_Start.Enabled := True;
                          BTN_Uninstall.Enabled := True;
                          BTN_Stop.Enabled := True;
                        end
                      else
                        begin
                          MessageBox( 0, 'La configuración fue guardada pero el replicador no podra ser ejecutado.', 'Mensaje de configuración', MB_ICONWARNING );

                          Empresas1.Enabled := False;
                          Correo1.Enabled := False;
                          Diasderecepcin1.Enabled := False;

                          BTN_Install.Enabled := False;
                          BTN_Start.Enabled := False;

                          if ( not isRunning( 'SyncService' ) ) and ( not isInstalled( 'SyncService' ) ) then
                            begin
                              BTN_Uninstall.Enabled := False;
                              BTN_Stop.Enabled := False;
                            end;
                        end;
                    end;
                end;
            end;
        end;



      BTN_ACCEPT.Caption := 'Aceptar';
      BTN_ACCEPT.Cursor := crDefault;
      BTN_ACCEPT.Enabled := True;
    end;
end;

procedure TT.BTN_CANCELClick(Sender: TObject);
begin
  Close;
end;





procedure TT.Empresas1Click(Sender: TObject);
begin
  Application.CreateForm( TTC, TC );
  TC.ShowModal;
  TC.Destroy;
end;

procedure TT.Correo1Click(Sender: TObject);
begin
  Application.CreateForm( TTM, TM );
  TM.ShowModal;
  TM.Destroy;
end;

procedure TT.Diasderecepcin1Click(Sender: TObject);
begin
  Application.CreateForm( TTR, TR );
  TR.ShowModal;
  TR.Destroy;
end;





{$REGION 'EVENTOS CLIC PARA SERVICIOS ( INSTALAR / DESINSTALAR / INICIAR / DETENER )'}

procedure TT.BTN_InstallClick(Sender: TObject);
begin
  if ( FileExists( 'Timer.ini' ) ) then
    begin
      if not isInstalled( 'SyncService' ) then
        begin
          if ( ShellExecute( 0, nil, 'Service.exe', '/INSTALL', nil, SW_SHOW ) <= 32 ) then
            begin
              MessageBox( 0, 'No fue posible instalar el servicio.', 'Mensaje de configuración', MB_ICONWARNING );
            end
          else
            begin
              MessageBox( 0, 'Servicio instalado satisfactoriamente.', 'Mensaje de configuración', MB_ICONINFORMATION );
            end;
        end
      else
        begin
          MessageBox( 0, 'El servicio ya se encuentra instalado.', 'Mensaje de configuración', MB_ICONWARNING );
        end;
    end
  else
    begin
      MessageBox( 0, 'El archivo ''Timer.ini'' no se encuentra en la carpeta del sistema.', 'Mensaje de configuración', MB_ICONWARNING );
    end;
end;

procedure TT.BTN_UninstallClick(Sender: TObject);
begin
  if isInstalled( 'SyncService' ) then
    begin
      if not isRunning( 'SyncService' ) then
        begin
          if ( ShellExecute( 0, nil, 'Service.exe', '/UNINSTALL', nil, SW_SHOW ) <= 32 ) then
            begin
              MessageBox( 0, 'No fue posible desinstalar el servicio.', 'Mensaje de configuración', MB_ICONWARNING );
            end
          else
            begin
              MessageBox( 0, 'Servicio desinstalado satisfactoriamente.', 'Mensaje de configuración', MB_ICONINFORMATION );
            end;
        end
      else
        begin
          MessageBox( 0, 'El servicio esta en ejecución primero hay que detenerlo.', 'Mensaje de configuración', MB_ICONWARNING );
        end;
    end
  else
    begin
      MessageBox( 0, 'El servicio no se encuentra instalado.', 'Mensaje de configuración', MB_ICONWARNING );
    end;
end;

procedure TT.BTN_StartClick(Sender: TObject);
begin
  if isInstalled( 'SyncService' ) then
    begin
      if not isRunning( 'SyncService' ) then
        begin
          StartSrv( 'SyncService' );
          MessageBox( 0, 'Servicio iniciado satisfactoriamente.', 'Mensaje de configuración', MB_ICONINFORMATION );
        end
      else
        begin
          MessageBox( 0, 'El servicio ya esta en ejecución', 'Mensaje de configuración', MB_ICONINFORMATION );
        end;
    end
  else
    begin
      MessageBox( 0, 'El servicio no se encuentra instalado.', 'Mensaje de configuración', MB_ICONWARNING );
    end;
end;

procedure TT.BTN_StopClick(Sender: TObject);
begin
  if isInstalled( 'SyncService' ) then
    begin
      if isRunning( 'SyncService' ) then
        begin
          StopService( 'SyncService' );
          MessageBox( 0, 'Servicio detenido satisfactoriamente.', 'Mensaje de configuración', MB_ICONINFORMATION );
        end
      else
        begin
          MessageBox( 0, 'El servicio no se encuentra en ejecución.', 'Mensaje de configuración', MB_ICONWARNING );
        end;
    end
  else
    begin
      MessageBox( 0, 'El servicio no se encuentra instalado.', 'Mensaje de configuración', MB_ICONWARNING );
    end;
end;

{$ENDREGION}

end.
