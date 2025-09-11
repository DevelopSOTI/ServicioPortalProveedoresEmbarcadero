unit Func_Catalogos;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf;

  // FUNCIONES PARA ACTUALIZAR LOS CATALOGOS (SUBIDA)
  Function ACTUALIZA_ALMACENES():Boolean;
  Function ACTUALIZA_MONEDAS():Boolean;
  Function ACTUALIZA_PROVEEDORES():Boolean;

  // FUNCIONES PARA ACTUALIZAR LOS CATALOGOS (BAJADA)
  Function UPDATE_LIBRES():Boolean;

implementation

uses
  Data, Func;


{$REGION 'ACTUALIZA_ALMACENES - ACTUALIZA LOS ALMACENES DE UNA EMPRESA ESPECIFICADA'}
Function ACTUALIZA_ALMACENES():Boolean;
  var
    Fmt :TFormatSettings;
    Almacen_ID, Nombre, Abreviado, Modif, Empresa_ID :string;
    ConnectionString, Command :string;
begin
  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat := 'hh:nn:ss';
  Fmt.TimeSeparator := ':';

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Almacenes')) then
    begin
      try
        ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;';
        ConnectionString := ConnectionString + 'UID=' + D.MYSQL_USER + ';';
        ConnectionString := ConnectionString + 'PORT=' + D.MYSQL_PORT + ';';
        ConnectionString := ConnectionString + 'DATABASE=' + D.MYSQL_DATA + ';';
        ConnectionString := ConnectionString + 'SERVER=' + D.MYSQL_SERV + ';';
        ConnectionString := ConnectionString + 'PASSWORD=' + D.MYSQL_PASS + ';';

        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := ConnectionString;
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Almacen.Close;
        D.JvCsvDataSet_Almacen.FileName := ExtractFilePath(ParamStr(0)) + '/Update/Almacenes';
        D.JvCsvDataSet_Almacen.Open;
        D.JvCsvDataSet_Almacen.First;
        while not D.JvCsvDataSet_Almacen.Eof do
          begin
            Application.ProcessMessages;
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Buscando almacen ' + D.JvCsvDataSet_Almacen.FieldByName('NOMBRE').AsString, '', '');

            try
              Almacen_ID := D.JvCsvDataSet_Almacen.FieldByName('ALMACEN_ID').AsString;
              Nombre := D.JvCsvDataSet_Almacen.FieldByName('NOMBRE').AsString;
              Abreviado := D.JvCsvDataSet_Almacen.FieldByName('NOMBRE_ABREV').AsString;
              Modif := FormatDateTime('YYYY-MM-DD HH:NN:SS', StrToDateTime(D.JvCsvDataSet_Almacen.FieldByName('FECHA_HORA_ULT_MODIF').AsString, Fmt));
              Empresa_ID := D.JvCsvDataSet_Almacen.FieldByName('EMPRESA_ID').AsString;

              D.ADOQueryActual.Active := False;
              D.ADOQueryActual.SQL.Clear;
              D.ADOQueryActual.SQL.Add('SELECT * FROM ALMACENES_MSP');
              D.ADOQueryActual.SQL.Add(' WHERE ALMACEN_ID_MSP = ' + Almacen_ID);
              D.ADOQueryActual.SQL.Add('   AND EMP_FK = ' + Empresa_ID);
              D.ADOQueryActual.Active := True;
              D.ADOQueryActual.First;

              if (D.ADOQueryActual.RecordCount = 0) then
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Registrando almacen ' + Nombre, '', '');

                  Command := 'INSERT INTO ALMACENES_MSP';
                  Command := Command + '(';
                  Command := Command + '  ALMACEN_ID_MSP,';
                  Command := Command + '  NOMBRE,';
                  Command := Command + '  NOMBRE_ABREV,';
                  Command := Command + '  FECHA_HORA_ULT_MODIF,';
                  Command := Command + '  EMP_FK';
                  Command := Command + ')';
                  Command := Command + 'VALUES';
                  Command := Command + '(';
                  Command := Command + '  ' + Almacen_ID + ',';
                  Command := Command + '  ' + QuotedStr(Nombre) + ',';
                  Command := Command + '  ' + QuotedStr(Abreviado) + ',';
                  Command := Command + '  ' + QuotedStr(Modif) + ',';
                  Command := Command + '  ' + Empresa_ID;
                  Command := Command + ')';
                end
              else
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando almacen ' + Nombre, '', '');

                  Command := 'UPDATE ALMACENES_MSP SET ';
                  Command := Command + '      NOMBRE = ' + QuotedStr(Nombre) + ',';
                  Command := Command + '      NOMBRE_ABREV = ' + QuotedStr(Abreviado) + ',';
                  Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(Modif) + ' ';
                  Command := Command + 'WHERE ALMACEN_ID_MSP = ' + Almacen_ID;
                  Command := Command + '  AND EMP_FK = ' + Empresa_ID;
                end;

              D.MySQL_Command.CommandText := Command;
              D.MySQL_Command.Execute;
            except
              on E : Exception do
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los almacenes');
                end;
            end;

            Inc(D.Position);
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

            D.JvCsvDataSet_Almacen.Next;
          end;

        Result := True;
        D.JvCsvDataSet_Almacen.Close;
      except
        on E : Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los almacenes');
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Almacenes'));
    end
  else
    begin
      Result := True;
    end;
end;
{$ENDREGION}

{$REGION 'ACTUALIZA_MONEDAS - ACTUALIZA LAS MONEDAS DE UNA EMPRESA ESPECIFICADA'}
Function ACTUALIZA_MONEDAS():Boolean;
  var
    Fmt :TFormatSettings;
    Moneda_ID, Nombre, Clave_Fiscal, Modif, Empresa_ID :string;
    ConnectionString, Command :string;
begin
  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat := 'hh:nn:ss';
  Fmt.TimeSeparator := ':';

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Monedas')) then
    begin
      try
        ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;';
        ConnectionString := ConnectionString + 'UID=' + D.MYSQL_USER + ';';
        ConnectionString := ConnectionString + 'PORT=' + D.MYSQL_PORT + ';';
        ConnectionString := ConnectionString + 'DATABASE=' + D.MYSQL_DATA + ';';
        ConnectionString := ConnectionString + 'SERVER=' + D.MYSQL_SERV + ';';
        ConnectionString := ConnectionString + 'PASSWORD=' + D.MYSQL_PASS + ';';

        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := ConnectionString;
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Moneda.Close;
        D.JvCsvDataSet_Moneda.FileName := ExtractFilePath(ParamStr(0)) + '/Update/Monedas';
        D.JvCsvDataSet_Moneda.Open;
        D.JvCsvDataSet_Moneda.First;
        while not D.JvCsvDataSet_Moneda.Eof do
          begin
            Application.ProcessMessages;
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Buscando moneda ' + D.JvCsvDataSet_Moneda.FieldByName('NOMBRE').AsString, '', '');

            try
              Moneda_ID := D.JvCsvDataSet_Moneda.FieldByName('MONEDA_ID').AsString;
              Nombre := D.JvCsvDataSet_Moneda.FieldByName('NOMBRE').AsString;
              Clave_Fiscal := D.JvCsvDataSet_Moneda.FieldByName('CLAVE_FISCAL').AsString;
              Modif := FormatDateTime('YYYY-MM-DD HH:NN:SS', StrToDateTime(D.JvCsvDataSet_Moneda.FieldByName('FECHA_HORA_ULT_MODIF').AsString, Fmt));
              Empresa_ID := D.JvCsvDataSet_Moneda.FieldByName('EMPRESA_ID').AsString;

              D.ADOQueryActual.Active := False;
              D.ADOQueryActual.SQL.Clear;
              D.ADOQueryActual.SQL.Add('SELECT * FROM MONEDAS_MSP');
              D.ADOQueryActual.SQL.Add(' WHERE MONEDA_ID = ' + Moneda_ID);
              D.ADOQueryActual.SQL.Add('   AND EMP_FK = ' + Empresa_ID);
              D.ADOQueryActual.Active := True;
              D.ADOQueryActual.First;

              if (D.ADOQueryActual.RecordCount = 0) then
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Registrando moneda ' + Nombre, '', '');

                  Command := 'INSERT INTO MONEDAS_MSP';
                  Command := Command + '(';
                  Command := Command + '  MONEDA_ID,';
                  Command := Command + '  NOMBRE,';
                  Command := Command + '  CLAVE_FISCAL,';
                  Command := Command + '  FECHA_HORA_ULT_MODIF,';
                  Command := Command + '  EMP_FK';
                  Command := Command + ')';
                  Command := Command + 'VALUES';
                  Command := Command + '(';
                  Command := Command + '  ' + Moneda_ID + ',';
                  Command := Command + '  ' + QuotedStr(Nombre) + ',';
                  Command := Command + '  ' + QuotedStr(Clave_Fiscal) + ',';
                  Command := Command + '  ' + QuotedStr(Modif) + ',';
                  Command := Command + '  ' + Empresa_ID;
                  Command := Command + ')';
                end
              else
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando moneda ' + Nombre, '', '');

                  Command := 'UPDATE MONEDAS_MSP SET ';
                  Command := Command + '      NOMBRE = ' + QuotedStr(Nombre) + ',';
                  Command := Command + '      CLAVE_FISCAL = ' + QuotedStr(Clave_Fiscal) + ',';
                  Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(Modif) + ' ';
                  Command := Command + 'WHERE MONEDA_ID = ' + Moneda_ID;
                  Command := Command + '  AND EMP_FK = ' + Empresa_ID;
                end;

              D.MySQL_Command.CommandText := Command;
              D.MySQL_Command.Execute;
            except
              on E : Exception do
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar las monedas');
                end;
            end;

            Inc(D.Position);
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

            D.JvCsvDataSet_Moneda.Next;
          end;

        Result := True;
        D.JvCsvDataSet_Moneda.Close;
      except
        on E : Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar las monedas');
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Monedas'));
    end
  else
    begin
      Result := True;
    end;
end;
{$ENDREGION}

{$REGION 'ACTUALIZA_PROVEEDORES - ACTUALIZA LOS PROVEEDORES DE UNA EMPRESA ESPECIFICADA'}
Function ACTUALIZA_PROVEEDORES():Boolean;
  var
    Fmt :TFormatSettings;
    prid, name, stat, clve, date, pctje_rechazo, referencia, rfc, s_recepcion, Empresa_ID :String;
    ConnectionString, Command :string;
begin
  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat := 'hh:nn:ss';
  Fmt.TimeSeparator := ':';

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Proveedores')) then
    begin
      try
        ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;';
        ConnectionString := ConnectionString + 'UID=' + D.MYSQL_USER + ';';
        ConnectionString := ConnectionString + 'PORT=' + D.MYSQL_PORT + ';';
        ConnectionString := ConnectionString + 'DATABASE=' + D.MYSQL_DATA + ';';
        ConnectionString := ConnectionString + 'SERVER=' + D.MYSQL_SERV + ';';
        ConnectionString := ConnectionString + 'PASSWORD=' + D.MYSQL_PASS + ';';

        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := ConnectionString;
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Proveedor.Close;
        D.JvCsvDataSet_Proveedor.FileName := ExtractFilePath(ParamStr(0)) + '/Update/Proveedores';
        D.JvCsvDataSet_Proveedor.Open;
        D.JvCsvDataSet_Proveedor.First;
        while not D.JvCsvDataSet_Proveedor.Eof do
          begin
            Application.ProcessMessages;
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Buscando proveedor ' + D.JvCsvDataSet_Proveedor.FieldByName('NOMBRE').AsString, '', '');

            try
              prid := D.JvCsvDataSet_Proveedor.FieldByName('PROVEEDOR_ID').AsString;
              name := D.JvCsvDataSet_Proveedor.FieldByName('NOMBRE').AsString;;
              stat := D.JvCsvDataSet_Proveedor.FieldByName('ESTATUS').AsString;;
              clve := D.JvCsvDataSet_Proveedor.FieldByName('CLAVE_PROV').AsString;;
              date := FormatDateTime('YYYY-MM-DD HH:NN:SS', StrToDateTime(D.JvCsvDataSet_Proveedor.FieldByName('FECHA_HORA_ULT_MODIF').AsString, Fmt));
              pctje_rechazo := D.JvCsvDataSet_Proveedor.FieldByName('PCTJE_RECHAZO').AsString;
              referencia := D.JvCsvDataSet_Proveedor.FieldByName('REFERENCIA').AsString;
              rfc := D.JvCsvDataSet_Proveedor.FieldByName('RFC_CURP').AsString;
              s_recepcion := D.JvCsvDataSet_Proveedor.FieldByName('PERMITIR_SIN_RECEPCION').AsString;
              Empresa_ID := D.JvCsvDataSet_Proveedor.FieldByName('EMPRESA_ID').AsString;

              D.ADOQueryActual.Active := False;
              D.ADOQueryActual.SQL.Clear;
              D.ADOQueryActual.SQL.Add('SELECT * FROM PROVEEDORES_MSP');
              D.ADOQueryActual.SQL.Add(' WHERE PROVEEDOR_ID_MSP = ' + prid);
              D.ADOQueryActual.SQL.Add('   AND EMP_FK = ' + Empresa_ID);
              D.ADOQueryActual.Active := True;
              D.ADOQueryActual.First;

              if (D.ADOQueryActual.RecordCount = 0) then
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Registrando proveedor ' + D.JvCsvDataSet_Proveedor.FieldByName('NOMBRE').AsString, '', '');

                  Command := 'INSERT INTO PROVEEDORES_MSP';
                  Command := Command + '(';
                  Command := Command + '  NOMBRE,';
                  Command := Command + '  ESTATUS,';
                  Command := Command + '  CLAVE_MSP,';
                  Command := Command + '  EMP_FK,';
                  Command := Command + '  FECHA_ULT_MODIF,';
                  Command := Command + '  PROVEEDOR_ID_MSP,';
                  Command := Command + '  RFC,';
                  Command := Command + '  PROV_PRIV,';
                  Command := Command + '  PCTJE_RECHAZO,';
                  Command := Command + '  REFERENCIA';
                  Command := Command + ')';
                  Command := Command + 'VALUES';
                  Command := Command + '(';
                  Command := Command + '  ' + QuotedStr(name) + ',';
                  Command := Command + '  ' + QuotedStr(stat) + ',';
                  Command := Command + '  ' + QuotedStr(clve) + ',';
                  Command := Command + '  ' + Empresa_ID + ',';
                  Command := Command + '  ' + QuotedStr(date) + ',';
                  Command := Command + '  ' + prid + ',';
                  Command := Command + '  ' + QuotedStr(rfc) + ',';
                  Command := Command + '  ' + QuotedStr(s_recepcion) + ',';
                  Command := Command + '  ' + QuotedStr(pctje_rechazo) + ',';
                  Command := Command + '  ' + QuotedStr(referencia);
                  Command := Command + ')';
                end
              else
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando proveedor ' + D.JvCsvDataSet_Proveedor.FieldByName('NOMBRE').AsString, '', '');

                  Command := 'UPDATE PROVEEDORES_MSP SET ';
                  Command := Command + '      NOMBRE = ' + QuotedStr(name) + ',';
                  Command := Command + '      ESTATUS = ' + QuotedStr(stat) + ',';
                  Command := Command + '      CLAVE_MSP = ' + QuotedStr(clve) + ',';
                  Command := Command + '      PROVEEDOR_ID_MSP = ' + prid + ',';
                  Command := Command + '      FECHA_ULT_MODIF = ' + QuotedStr(date) + ',';
                  Command := Command + '      RFC = ' + QuotedStr(rfc) + ',';
                  Command := Command + '      PROV_PRIV = ' + QuotedStr(s_recepcion) + ',';
                  Command := Command + '      PCTJE_RECHAZO = ' + QuotedStr(pctje_rechazo) + ', ';
                  Command := Command + '      REFERENCIA = ' + QuotedStr(referencia) + ' ';
                  Command := Command + 'WHERE PROVEEDOR_ID_MSP = ' + prid;
                  Command := Command + '  AND EMP_FK = ' + Empresa_ID;
                end;

              D.MySQL_Command.CommandText := Command;
              D.MySQL_Command.Execute;
            except
              on E : Exception do
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los proveedores');
                end;
            end;

            Inc(D.Position);
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

            D.JvCsvDataSet_Proveedor.Next;
          end;

        Result := True;
        D.JvCsvDataSet_Proveedor.Close;
      except
        on E : Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los proveedores');
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Proveedores'));
    end
  else
    begin
      Result := True;
    end;
end;
{$ENDREGION}



{$REGION 'UPDATE_LIBRES - ACTUALIZA LOS DATOS LIBRES DE LOS PROVEEDORES EN MICROSIP SEGUN EL PORTAL'}
Function UPDATE_LIBRES():Boolean;
  var
    proveedor_id, cuenta, sucursal, clabe, banco_nombre, mail, referencia, nombre, Empresa_ID :String;
begin
  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Revisando datos libres de los proveedores en el portal', '', '');

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Libres')) then
    begin
      try
        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Libre.Close;
        D.JvCsvDataSet_Libre.FileName := ExtractFilePath(ParamStr(0)) + '/Update/Libres';
        D.JvCsvDataSet_Libre.Open;
        D.JvCsvDataSet_Libre.First;
        while not D.JvCsvDataSet_Libre.Eof do
          begin
            Application.ProcessMessages;
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando datos particulares del proveedor ' + D.JvCsvDataSet_Libre.FieldByName('NOMBRE').AsString, '', '');

            proveedor_id := D.JvCsvDataSet_Libre.FieldByName('PROVEEDOR_ID_MSP').AsString;
            cuenta := D.JvCsvDataSet_Libre.FieldByName('CUENTA').AsString;
            sucursal := D.JvCsvDataSet_Libre.FieldByName('SUCURSAL').AsString;
            clabe := D.JvCsvDataSet_Libre.FieldByName('CLABE').AsString;
            banco_nombre := D.JvCsvDataSet_Libre.FieldByName('BANCO_NOMBRE').AsString;
            mail := D.JvCsvDataSet_Libre.FieldByName('MAIL').AsString;
            referencia := D.JvCsvDataSet_Libre.FieldByName('REFERENCIA').AsString;
            nombre := D.JvCsvDataSet_Libre.FieldByName('NOMBRE').AsString;
            Empresa_ID := D.JvCsvDataSet_Libre.FieldByName('EMPRESA_ID').AsString;

            try
              // CONEXIÓN MICROSIP (SI ES QUE CAMBIA DE EMPRESA)
              if (D.Conexion_Microsip.DatabaseName <> (D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Libre.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB')) then
                begin
                  D.Conexion_Microsip.Connected := False;
                  D.Conexion_Microsip.DatabaseName := D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Libre.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB';
                  D.Conexion_Microsip.Connected := True;
                  D.Transaction_Microsip.Active := True;
                end;

              D.IBQueryMicrosip.Active := False;
              D.IBQueryMicrosip.SQL.Clear;
              // D.IBQueryMicrosip.SQL.Add('UPDATE LIBRES_PROVEEDOR SET CTA_TRANSFERENCIA_ELECTRONICA = ' + QuotedStr(cuenta) + ', SUCURSAL = ' + QuotedStr(sucursal) + ', CLABE = ' + QuotedStr(clabe) + ', BANCO = ' + QuotedStr(banco_nombre) + ', CORREO_CXC = ' + QuotedStr(mail) + ', REFERENCIA = ' + QuotedStr(referencia) + ' WHERE PROVEEDOR_ID = ' + proveedor_id);
              D.IBQueryMicrosip.SQL.Add('UPDATE LIBRES_PROVEEDOR SET ');
              D.IBQueryMicrosip.SQL.Add('       CTA_TRANSFERENCIA_ELECTRONICA = ' + QuotedStr(cuenta) + ', ');
              D.IBQueryMicrosip.SQL.Add('       SUCURSAL = ' + QuotedStr(sucursal) + ', ');
              D.IBQueryMicrosip.SQL.Add('       CLABE = ' + QuotedStr(clabe) + ', ');
              D.IBQueryMicrosip.SQL.Add('       BANCO = ' + QuotedStr(banco_nombre) + ', ');
              D.IBQueryMicrosip.SQL.Add('       CORREO_CXC = ' + QuotedStr(mail) + ', ');
              D.IBQueryMicrosip.SQL.Add('       REFERENCIA = ' + QuotedStr(referencia) + ' ');
              D.IBQueryMicrosip.SQL.Add(' WHERE PROVEEDOR_ID = ' + proveedor_id);
              D.IBQueryMicrosip.Active := True;

              D.Transaction_Microsip.Commit;
            except
              on E:Exception do
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los datos libres del proveedor ' + nombre);
                  D.Transaction_Microsip.RollbackRetaining;
                end;
            end;

            Inc(D.Position);
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

            D.JvCsvDataSet_Libre.Next;
          end;

        // NOS DESCONECTAMOS
        D.Transaction_Microsip.Active := False;
        D.Conexion_Microsip.Connected := False;

        Result := True;
        D.JvCsvDataSet_Libre.Close;
      except
        on E : Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al revisar los datos libres de los proveedores');
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Libres'));
    end
  else
    begin
      Result := True;
    end;
end;
{$ENDREGION}

end.
