unit Func;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf;

  // FUNCIONES PARA CONTROL DEL SERVICIOS
  Function isInstalled(Nombre: string): Boolean;
  Function isRunning(Nombre: string): Boolean;
  procedure StopService(Nombre: string);
  procedure StartSrv(Nombre: string);

  // FUNCIONES GENERALES
  procedure EVENT_LOG(Max, Row, Title, Last, Error :string; Incluir :Boolean = False; Process :string = 'I');
  procedure GET_APLICA_FACTURAS();
  procedure GET_LAST_UPDATE();

  Function GET_PLAZO(id :string):string;
  Function SIGUIENTE_FOLIO(SERIE :string):string;
  Function FORMAT_FOLIO(serie, folio :string):string;
  Function GET_USO_CLAVE(uso_cfdi :string):string;

  // FUNCIONES PARA ENVIO DE CORREO
  procedure PROCESO_ENVIAR(proveedor_id :Integer; FolioFac, FechaFac, FechaProv :string);

  // FUNCIONES PARA ACTUALIZAR LOS REGISTROS (PRINCIPAL)
  Function INSERT_UPDATE_EMPRESAS():Boolean;
  Function INSERT_UPDATE():Boolean;

implementation

uses
  Data, Form, Main, Func_Catalogos, Func_Recepciones, Func_Creditos, Func_Facturas_3_2, Func_Facturas_3_3;

{$REGION 'FUNCIONES PARA EL CONTROL DEL SERVICIO (INSTALACIÓN, INICIO, DETENER Y DESINSTALACIÓN)'}

{$REGION 'isInstalled - VERIFICA SI EL SERVICIO ESTA INSTALADO'}
Function isInstalled(Nombre :string): Boolean;
  var
    ServiceControlManager :SC_HANDLE;
    Service :SC_HANDLE;
    // ServiceStatus :SERVICE_STATUS;
begin
  Result := False;
  ServiceControlManager := OpenSCManager( nil, nil, SC_MANAGER_CONNECT );

  if ServiceControlManager <> 0 then
    begin
      Service := OpenService( ServiceControlManager, PChar( Nombre ), GENERIC_READ );

      if Service <> 0 then
        begin
          Result := True;
          CloseServiceHandle( Service );
        end;

      CloseServiceHandle( ServiceControlManager );
    end;
end;
{$ENDREGION}

{$REGION 'isRunning - VERIFICA SI EL SERVICIO ESTA CORRIENDO'}
Function isRunning(Nombre :string): Boolean;
  var
    ServiceControlManager :SC_HANDLE;
    Service :SC_HANDLE;
    ServiceStatus :SERVICE_STATUS;
begin
  Result := False;
  ServiceControlManager := OpenSCManager( nil, nil, SC_MANAGER_CONNECT );

  if ServiceControlManager <> 0 then
    begin
      Service := OpenService( ServiceControlManager, PChar( Nombre ), GENERIC_READ );

      if Service <> 0 then
        begin
          if QueryServiceStatus( Service, ServiceStatus ) then
            begin
              Result := ServiceStatus.dwCurrentState = SERVICE_RUNNING;
            end;

          CloseServiceHandle( Service );
        end;

      CloseServiceHandle( ServiceControlManager );
    end;
end;
{$ENDREGION}

{$REGION 'StopService - DETIENE EL SERVICIO'}
procedure StopService(Nombre :string);
  var
    ServiceControlManager :SC_HANDLE;
    Service :SC_HANDLE;
    ServiceStatus :SERVICE_STATUS;
begin
  ServiceControlManager := OpenSCManager( nil, nil, SC_MANAGER_CONNECT );

  if ServiceControlManager <> 0 then
    begin
      Service := OpenService( ServiceControlManager, PChar( Nombre ), SERVICE_ALL_ACCESS );

      if Service <> 0 then
        begin
          if QueryServiceStatus( Service, ServiceStatus ) then
            begin
              if ServiceStatus.dwCurrentState <> SERVICE_STOPPED  then
                begin
                  ControlService( Service, SERVICE_CONTROL_STOP, ServiceStatus );
                end;
            end;

          CloseServiceHandle( Service );
        end;

      CloseServiceHandle( ServiceControlManager );
    end;
end;
{$ENDREGION}

{$REGION 'StartSrv - INICIA EL SERVICIO'}
procedure StartSrv(Nombre :string);
  var
    ServiceControlManager :SC_HANDLE;
    Service :SC_HANDLE;
    ServiceStatus :SERVICE_STATUS;
    Argv :PChar;
begin
  ServiceControlManager := OpenSCManager( nil, nil, SC_MANAGER_CONNECT );

  if ServiceControlManager <> 0 then
    begin
      Service := OpenService( ServiceControlManager, PChar( Nombre ), SERVICE_ALL_ACCESS );

      if Service <> 0 then
        begin
          if QueryServiceStatus( Service, ServiceStatus ) then
            begin
              if ServiceStatus.dwCurrentState <> SERVICE_RUNNING  then
                begin
                  Argv := nil;
                  StartService( Service, 0, Argv );
                end;
            end;

          CloseServiceHandle( Service );
        end;

      CloseServiceHandle( ServiceControlManager );
    end;
end;
{$ENDREGION}

{$ENDREGION}



{$REGION 'FUNCIONES GENERALES'}

{$REGION 'EVENT_LOG - CREA LOS ARCHIVOS QUE LEE LA PANTALLA DE ESCRITORIO'}
procedure EVENT_LOG(Max, Row, Title, Last, Error :string; Incluir :Boolean = False; Process :string = 'I');
  var
    RutaFichero :string;
    EventFile :TStringList;
begin
  if (D.MODE_APPLI = 'F') then
    begin
      F.ProgressBar.Position := StrToInt(Row);
      F.ProgressBar.Update;

      F.Panel.Caption := Title;
      if (Error <> '') then
        begin
          F.Memo.Lines.Add(Error);
        end;
      F.Update;
    end
  else
    begin
      RutaFichero := ExtractFilePath(ParamStr(0)) + 'EventLog\EventLog.ini';
      EventFile := TStringList.Create;

      if (Error <> '') then
        begin
          // M.EventLog.LogMessage(Error);
        end;

      EventFile.Add(Max); // MAXIMO DEL PROGRESSBAR
      EventFile.Add(Row); // POSICIÓN EN EL PROGRESS ACTUAL
      EventFile.Add(Title); // PROCESO ACTUAL
      EventFile.Add(Last); // ULTIMA FECHA
      EventFile.Add(Error); // ULTIMO ERROR

      if (Incluir = True) then
        begin
          EventFile.Add(Process);
        end;

      try
        EventFile.SaveToFile(RutaFichero);
        Sleep(30);
      except
      end;
    end;
end;
{$ENDREGION}

{$REGION 'GET_APLICA_FACTURAS - OBTENEMOS EL PARAMETRO QUE INDICA SI LAS FACTURAS SE INSERTAN AUTOMATICA O MANUALMENTE'}
procedure GET_APLICA_FACTURAS();
begin
  try
    D.ADOQueryParametros.Active := False;
    D.ADOQueryParametros.SQL.Clear;
    D.ADOQueryParametros.SQL.Add('SELECT PARAM_VALOR FROM PARAMETROS WHERE PARAM_CLAVE = ''APLICA_DIR''');
    D.ADOQueryParametros.Active := True;

    if (D.ADOQueryParametros.RecordCount = 1) then
      begin
        if (D.ADOQueryParametros.FieldByName('PARAM_VALOR').AsString = 'TRUE') then
          begin
            D.AplicaFacturas := True;
          end
        else
          begin
            D.AplicaFacturas := False;
          end;
      end
    else
      begin
        D.MySQL_Command.CommandText := 'INSERT INTO PARAMETROS(PARAM_CLAVE, PARAM_VALOR) VALUES(''APLICA_DIR'', ''FALSE'')';
        D.MySQL_Command.Execute;

        D.AplicaFacturas := False; // SI NO HAY RENGLONES RETORNAMOS FALSE Y AGREGA EL PARAMETRO CON VALOR PREDETERMINADO EN FALSO
      end;

    D.ADOQueryParametros.Active := False;
  except
    on E:Exception do
      begin
        D.ADOQueryParametros.Active := False;
        D.AplicaFacturas := False; // SI HAY ALGUN ERROR RETORNAMOS FALSE

        EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al obtener el metodo de inserción de facturas del portal.');
      end;
  end;
end;
{$ENDREGION}

{$REGION 'GET_LAST_UPDATE - OBTENEMOS LA ULTIMA FECHA DE ACTUALIZACIÓN, MENOS UN DIA'}
procedure GET_LAST_UPDATE();
  var
    Fmt :TFormatSettings;
    dt :TDateTime;
begin
  try
    Fmt.ShortDateFormat := 'dd/mm/yyyy';
    Fmt.DateSeparator := '/';
    Fmt.LongTimeFormat :='hh:nn:ss';
    Fmt.TimeSeparator  :=':';

    D.ADOQueryParametros.Active := False;
    D.ADOQueryParametros.SQL.Clear;
    D.ADOQueryParametros.SQL.Add('SELECT PARAM_VALOR FROM PARAMETROS WHERE PARAM_CLAVE = ''LAST_UPDATE''');
    D.ADOQueryParametros.Active := True;

    if (D.ADOQueryParametros.RecordCount = 1) then
      begin
        if (D.ADOQueryParametros.FieldByName('PARAM_VALOR').AsString <> '') then
          begin
            dt := StrToDateTime(D.ADOQueryParametros.FieldByName('PARAM_VALOR').AsString, Fmt);
            // D.LastUpdate_Text := FormatDateTime('dd/mm/yyyy hh:nn:ss', IncDay(dt, -1));
            D.LastUpdate_Text := FormatDateTime('dd/mm/yyyy hh:nn:ss', dt);
          end;
      end
    else
      begin
        D.MySQL_Command.CommandText := 'INSERT INTO PARAMETROS(PARAM_CLAVE, PARAM_VALOR) VALUES(''LAST_UPDATE'', null)';
        D.MySQL_Command.Execute;

        D.LastUpdate_Text := ''; // SI NO HAY RENGLONES DEVOLVEMOS EN BLANCO LA FECHA DE ULTIMA ACTUALIZACIÓN
      end;

    D.ADOQueryParametros.Active := False;
  except
    on E:Exception do
      begin
        D.ADOQueryParametros.Active := False;
        D.LastUpdate_Text := ''; // SI MARCA ERROR DEVOLVEMOS EN BLANCO LA FECHA DE ULTIMA ACTUALIZACIÓN

        EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al obtener la fecha de la ultima actualización.');
      end;
  end;
end;
{$ENDREGION}

{$REGION 'GET_PLAZO - FUNCIÓN QUE RETORNA UN PLAZO DE MICROSIP'}
Function GET_PLAZO(ID :string):string;
begin
  try
    D.IBQueryDetalle.Active := False;
    D.IBQueryDetalle.SQL.Clear;
    D.IBQueryDetalle.SQL.Add('SELECT P.dias_plazo FROM doctos_cm D');
    D.IBQueryDetalle.SQL.Add(' INNER JOIN condiciones_pago_cp C ON ( C.cond_pago_id = D.cond_pago_id )');
    D.IBQueryDetalle.SQL.Add(' INNER JOIN plazos_cond_pag_cp P ON ( P.cond_pago_id = C.cond_pago_id )');
    D.IBQueryDetalle.SQL.Add(' WHERE D.docto_cm_id = ' + ID);
    D.IBQueryDetalle.Active := True;

    Result := D.IBQueryDetalle.FieldByName('DIAS_PLAZO').AsString;
  except
    on E:Exception do
      begin
        EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al buscar el plazo del documento ' + ID + '.');

        Result := '0';
      end;
  end;

  D.IBQueryDetalle.Active := False;
end;
{$ENDREGION}

{$REGION 'SIGUIENTE_FOLIO - FUNCIÓN QUE DEVUELVE EL SIGUIENTE FOLIO A ASIGNAR DE UNA SERIE INDICADA'}
Function SIGUIENTE_FOLIO(SERIE :string):string;
  var
    VINT, LONG :Integer;
    VCAD :string;
begin
  // LEE EL ULTIMO FOLIO DE LA SERIE INDICADA
  D.IBQueryDetalle.Active := False;
  D.IBQueryDetalle.SQL.Clear;
  D.IBQueryDetalle.SQL.Add('SELECT * FROM folios_compras WHERE serie = ''' + SERIE + '''');
  D.IBQueryDetalle.Active := True;

  VINT := D.IBQueryDetalle.FieldByName('CONSECUTIVO').AsInteger;
  Inc(VINT);
  VCAD := IntToStr(VINT);
  // Inc(VINT);
  LONG := Length(VCAD);

  // ACTUALIZA EL NUEVO FOLIO
  D.IBQueryDetalle.Active := False;
  D.IBQueryDetalle.SQL.Clear;
  D.IBQueryDetalle.SQL.Add('UPDATE folios_compras SET CONSECUTIVO = ' + IntToStr(VINT) + 'WHERE serie = ''WEB''');
  D.IBQueryDetalle.Active := True;

  case LONG of
    1:  VCAD := '00000' + VCAD;
    2:  VCAD := '0000' + VCAD;
    3:  VCAD := '000' + VCAD;
    4:  VCAD := '00' + VCAD;
    5:  VCAD := '0' + VCAD;
    6:  VCAD := VCAD;
  end;

  Result := SERIE + VCAD;
end;
{$ENDREGION}

{$REGION 'FORMAT_FOLIO - FUNCIÓN QUE LE DA FORMATO DE 9 CARACTERES A UN FOLIO'}
Function FORMAT_FOLIO(Serie, Folio :string):string;
  var
    Formato :string;
begin
  Formato := Serie;
  while (Length(Formato) + Length(Folio) < 9) do
    begin
      Formato := Formato + '0';
    end;
  Formato := Formato + Folio;

  Result := Formato;
end;
{$ENDREGION}

Function GET_USO_CLAVE(uso_cfdi :string):string;
  var
    position :Integer;
begin
  Result := '';

  if Length(uso_cfdi) > 0 then
    begin
      position := Pos('-', uso_cfdi) - 1;
      Result := Trim(Copy(uso_cfdi, 0, position));
    end;
end;

{$ENDREGION}



{$REGION 'FUNCIONES PARA ENVIO DE CORREO'}

{$REGION 'EnviarMensaje - FUNCIÓN PARA ENVIAR CORREOS'}
procedure EnviarMensaje(sUsuario, sClave, sHost, sPort, sAsunto, sDestino, sMensaje: string);
  var
    SMTP :TIdSMTP;
    Mensaje :TIdMessage;
begin
  // Creamos el componente de conexión con el servidor
  SMTP := TIdSMTP.Create(nil);
  SMTP.Username := sUsuario;
  SMTP.Password := sClave;
  SMTP.Host := sHost;
  SMTP.Port := StrToInt(sPort);
  // SMTP.AuthenticationType := atLogin; // SOLO EN DELPHI 7

  // Creamos el contenido del mensaje
  Mensaje := TIdMessage.Create(nil);
  Mensaje.Clear;
  Mensaje.From.Name := sUsuario;
  Mensaje.From.Address := sDestino;
  Mensaje.Subject := sAsunto;
  Mensaje.Body.Text := sMensaje;
  Mensaje.Recipients.Add;
  Mensaje.Recipients.Items[0].Address := sDestino;

  // Conectamos con el servidor SMTP
  try
    SMTP.Connect;
  except
    raise Exception.Create('No fue posible conectar con el servidor.');
  end;

  // Si ha conectado enviamos el mensaje y desconectamos
  if SMTP.Connected then
    begin
      try
        SMTP.Send(Mensaje);
      except
        raise Exception.Create('No fue posible enviar el mensaje.');
      end;

      try
        SMTP.Disconnect;
      except
        raise Exception.Create('Hubo un error al desconectar del servidor.');
      end;
    end;

  FreeAndNil(Mensaje);
  FreeAndNil(SMTP);
end;
{$ENDREGION}

{$REGION 'PROCESO_ENVIAR - PROCESO QUE LLAMA A LA FUNCIÓN PARA ENVIAR CORREO'}
procedure PROCESO_ENVIAR(proveedor_id :Integer; FolioFac, FechaFac, FechaProv :string);
  var
    EMAIL, EMAIL1, NOMBRE, CORREO, PASS, SERVIDOR, PUERTO, MENSAJE :string;
begin
  // BUSCA INFORMACION EN MICROSIP
  D.IBQueryDetalle.Active := False;
  D.IBQueryDetalle.SQL.Clear;
  D.IBQueryDetalle.SQL.Add('SELECT * FROM proveedores p ');
  D.IBQueryDetalle.SQL.Add(' INNER JOIN libres_proveedor l ON (p.proveedor_id = l.proveedor_id) ');
  D.IBQueryDetalle.SQL.Add(' WHERE p.proveedor_id = ' + IntToStr(PROVEEDOR_ID));
  D.IBQueryDetalle.Active := True;

  EMAIL := D.IBQueryDetalle.FieldByName('EMAIL').AsString;
  EMAIL1 := D.IBQueryDetalle.FieldByName('CORREO_CXC').AsString;
  NOMBRE := D.IBQueryDetalle.FieldByName('NOMBRE').AsString;

  // BUSCA LOS DATOS DEL CORREO CONFIGURADO EN EL PORTAL
  D.ADOQueryMySQL.Active := False;
  D.ADOQueryMySQL.SQL.Clear;
  D.ADOQueryMySQL.SQL.ADD('SELECT * FROM MAIL');
  D.ADOQueryMySQL.Active := True;

  CORREO := D.ADOQueryMySQL.FieldByName('MAIL_FROM').AsString;
  PASS := D.ADOQueryMySQL.FieldByName('MAIL_PASS').AsString;
  SERVIDOR := D.ADOQueryMySQL.FieldByName('MAIL_SMTP').AsString;
  PUERTO := D.ADOQueryMySQL.FieldByName('MAIL_PORT').AsString;

  // ARMAMOS EL MENSAJE DEL CORREO
  MENSAJE := 'Estimado proveedor' + #13#13;
  MENSAJE := MENSAJE + 'Le notificamos que la factura ' + FolioFac + ' con fecha ' + FechaFac + ' fue recibida y paso a pendiente de pago, la fecha estimada de pago seria el dia ' + FechaProv + #13;
  MENSAJE := MENSAJE + 'Favor de verificar el estatus de la factura en el portal de proveedores.';

  // INTENTAMOS ENVIAR EL CORREO
  // if (EMAIL1 = '') then
  if (EMAIL1 <> '') then
    begin
      // EnviarMensaje(CORREO, PASS, SERVIDOR, PUERTO, 'Contra recibo electronico', EMAIL, MENSAJE)
      EnviarMensaje(CORREO, PASS, SERVIDOR, PUERTO, 'Contra recibo electronico', EMAIL1, MENSAJE)
    end
  else
    begin
      if (EMAIL <> '') then
        begin
          // EnviarMensaje(CORREO, PASS, SERVIDOR, PUERTO, 'Contra recibo electronico', EMAIL1, MENSAJE);
          EnviarMensaje(CORREO, PASS, SERVIDOR, PUERTO, 'Contra recibo electronico', EMAIL, MENSAJE);
        end
      else
        begin
          // M.EventLog.LogMessage('No hay ningun correo registrado con el proveedor ' + NOMBRE);
        end;
    end;
end;
{$ENDREGION}

{$ENDREGION}





{$REGION 'INSERT_UPDATE_EMPRESAS'}
Function INSERT_UPDATE_EMPRESAS():Boolean;
  var
    Fmt :TFormatSettings;
    Empresa_ID, Nombre, Fecha, Long, RFC :String;
begin
  EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando empresas', '', '');

  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat := 'hh:nn:ss';
  Fmt.TimeSeparator := ':';

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Empresas')) then
    begin
      try
        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Empresa.Close;
        D.JvCsvDataSet_Empresa.FileName := ExtractFilePath(ParamStr(0)) + '/Update/Empresas';
        D.JvCsvDataSet_Empresa.Open;
        D.JvCsvDataSet_Empresa.First;
        while not D.JvCsvDataSet_Empresa.Eof do
          begin
            Application.ProcessMessages;
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Buscando empresa ' + D.JvCsvDataSet_Empresa.FieldByName('NOMBRE_CORTO').AsString, '', '');

            Empresa_ID := D.JvCsvDataSet_Empresa.FieldByName('EMPRESA_ID').AsString;
            Nombre := D.JvCsvDataSet_Empresa.FieldByName('NOMBRE_CORTO').AsString;
            Fecha := FormatDateTime('YYYY-MM-DD HH:MM:SS', StrToDateTime(D.JvCsvDataSet_Empresa.FieldByName('FECHA_HORA_ULT_MODIF').AsString, Fmt));
            Long := D.JvCsvDataSet_Empresa.FieldByName('NOMBRE').AsString;
            RFC := D.JvCsvDataSet_Empresa.FieldByName('RFC').AsString;

            try
              D.MySQL_Command.CommandText := '';

              D.ADOQueryActual.Active := False;
              D.ADOQueryActual.SQL.Clear;
              D.ADOQueryActual.SQL.Add('SELECT * FROM EMPRESAS_MSP WHERE EMP_ID_MSP = ' + Empresa_ID);
              D.ADOQueryActual.Active := True;
              D.ADOQueryActual.First;

              if (D.ADOQueryActual.RecordCount = 0) then
                begin
                  EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Registrando empresa ' + Nombre, '', '');

                  D.MySQL_Command.CommandText := 'INSERT INTO EMPRESAS_MSP( EMP_ID_MSP, EMP_NOMBRE, EMP_FECHA_ULT_MODIF, EMP_NOMBRE_LARGO, EMP_RFC, EMP_ESTATUS ) ';
                  D.MySQL_Command.CommandText := D.MySQL_Command.CommandText + 'VALUES ( ' + Empresa_ID + ', ' + QuotedStr( Nombre ) + ', ' + QuotedStr( Fecha ) + ', ' + QuotedStr( Long ) + ', ' + QuotedStr( RFC ) + ', ''Bloqueada'' )';
                  D.MySQL_Command.Execute;
                end
              else
                begin
                  EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando empresa ' + Nombre, '', '');

                  D.MySQL_Command.CommandText := 'UPDATE EMPRESAS_MSP SET EMP_NOMBRE = ' + QuotedStr( Nombre ) + ', EMP_FECHA_ULT_MODIF = ' + QuotedStr( Fecha ) + ', EMP_NOMBRE_LARGO = ' + QuotedStr( Long ) + ', EMP_RFC = ' + QuotedStr( RFC ) + ' WHERE EMP_ID_MSP = ' + Empresa_ID;
                  D.MySQL_Command.Execute;
                end;
            except
              on E : Exception do
                begin
                  EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar las empresas');
                end;
            end;

            Inc(D.Position);
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

            D.JvCsvDataSet_Empresa.Next;
          end;

        Result := True;
        D.JvCsvDataSet_Empresa.Close;
      except
        on E : Exception do
          begin
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar las empresas');
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Empresas'));
    end
  else
    begin
      Result := True;
    end;
end;
{$ENDREGION}

{$REGION 'INSERT_UPDATE'}
Function INSERT_UPDATE():Boolean;
begin
  Result := False; // INICIALIZA EN FALSO POR SI EN ALGUNO DE LOS PROCESOS FALLA RETORNE FALSO LA FUNCIÓN

  {$REGION 'Func_Catalogos'}
  if (Func_Catalogos.ACTUALIZA_ALMACENES = False) then
    begin
      Exit;
    end;

  if (Func_Catalogos.ACTUALIZA_MONEDAS = False) then
    begin
      Exit;
    end;

  if (Func_Catalogos.ACTUALIZA_PROVEEDORES = False) then
    begin
      Exit;
    end;

  // ESTE YA NO SE ACTIVA
  // UPDATE_FACTURAS(Empresa_ID);

  if (Func_Catalogos.UPDATE_LIBRES = False) then
    begin
      Exit;
    end;
  {$ENDREGION}

  if (ACTUALIZA_RECEPCIONES = False) then
    begin
      Exit;
    end;

  { if (ACTUALIZA_CREDITOS = False) then
    begin
      Exit;
    end; }

  if (D.AplicaFacturas = True) then
    begin
      if (SELECT_FACTURAS_APLICAR_33 = False) then
        begin
          Exit;
        end;
    end;

  Result := True; // SI EN TODOS LOS PROCESOS SE DEVOLVIO VERDADERO RETORNA UN VERDADERO LA FUNCIÓN
end;
{$ENDREGION}

end.
