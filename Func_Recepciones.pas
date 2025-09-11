unit Func_Recepciones;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf;

  // F_ACTUALIZA_REGISTROS - FUNCIONES PARA ACTUALIZAR LOS REGISTROS (SUBIDA)
  Function ACTUALIZA_RECEPCIONES():Boolean;

implementation

uses
  Data, Func;


{$REGION 'ACTUALIZA_RECEPCIONES_DET - ACTUALIZA EL DETALLE DE LA RECEPCIÓN ESPECIFICADA'}
procedure ACTUALIZA_RECEPCIONES_DET( docto_id_old, folio, empresa_id :String );
  var
    docto_det, docto, nombre, unidades, precio_unitario, pctje_dscto, precio_total, notas, posicion :string;
    Command :string;

    UnicodeStr: UnicodeString;
    UTF8Str: RawByteString;
begin
  if (docto_id_old <> '') then
    begin
      try
        Command := 'DELETE FROM RECEPCIONES_DET ';
        Command := Command + 'WHERE DOCTO_CM_ID_MSP = ' + docto_id_old + ' ';
        Command := Command + '  AND EMP_FK = ' + empresa_id;

        D.MySQL_Command.CommandText := Command;
        D.MySQL_Command.Execute;
      except
        on E : Exception do
          begin
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los detalles de las recepciones');
          end;
      end;
    end;

  try
    D.IBQueryDetalle.Active := False;
    D.IBQueryDetalle.SQL.Clear;
    D.IBQueryDetalle.SQL.Add('SELECT dcd.docto_cm_det_id, dcd.docto_cm_id, art.nombre, dcd.unidades, dcd.precio_unitario, dcd.pctje_dscto, dcd.precio_total_neto, dcd.notas, dcd.posicion FROM doctos_cm_det dcd');
    D.IBQueryDetalle.SQL.Add('  JOIN doctos_cm dce ON ( dcd.docto_cm_id = dce.docto_cm_id )');
    D.IBQueryDetalle.SQL.Add('  JOIN articulos art ON ( dcd.articulo_id = art.articulo_id )');
    D.IBQueryDetalle.SQL.Add(' WHERE dce.tipo_docto = ''R'' AND dce.estatus = ''P'' AND dce.fecha > ''31.12.2014'' AND dce.folio = ''' + folio + '''');
    D.IBQueryDetalle.Active := True;
    D.IBQueryDetalle.First;
    while not D.IBQueryDetalle.Eof do
      begin
        Application.ProcessMessages;

        try
          docto_det := D.IBQueryDetalle.FieldByName('DOCTO_CM_DET_ID').AsString;
          docto := D.IBQueryDetalle.FieldByName('DOCTO_CM_ID').AsString;

          nombre := D.IBQueryDetalle.FieldByName('NOMBRE').AsString;
          nombre := StringReplace(nombre, '''', '''''', [rfReplaceAll]);
          nombre := StringReplace(nombre, '\', '\\', [rfReplaceAll]);
          // nombre := StringReplace(nombre, '”', '"', [rfReplaceAll]);
          // nombre := StringReplace(nombre, '?', '"', [rfReplaceAll]);
          nombre := StringReplace(nombre, '\xC2\x94', '"', [rfReplaceAll]);

          UTF8Str := UTF8Encode(nombre);
          SetCodePage(UTF8Str, 0, False);
          // UnicodeStr := UTF8Str;
          nombre := UTF8Str;

          unidades := StringReplace(FormatFloat('#.00', D.IBQueryDetalle.FieldByName('UNIDADES').AsFloat), ',', '.', [rfReplaceAll]);
          precio_unitario := StringReplace(FormatFloat('#.000000', D.IBQueryDetalle.FieldByName('PRECIO_UNITARIO').AsFloat), ',', '.', [rfReplaceAll]);
          pctje_dscto := StringReplace(FormatFloat('#.00', D.IBQueryDetalle.FieldByName('PCTJE_DSCTO').AsFloat), ',', '.', [rfReplaceAll]);
          precio_total := StringReplace(FormatFloat('#.000000', D.IBQueryDetalle.FieldByName('PRECIO_TOTAL_NETO').AsFloat), ',', '.', [rfReplaceAll]);
          notas := D.IBQueryDetalle.FieldByName('NOTAS').AsString;
          notas := StringReplace(notas, '''', '''''', [rfReplaceAll]);
          notas := StringReplace(notas, '\', '\\', [rfReplaceAll]);
          notas := StringReplace(notas, '--', '-', [rfReplaceAll]);
          posicion := D.IBQueryDetalle.FieldByName('POSICION').AsString;

          Command := 'INSERT INTO RECEPCIONES_DET';
          Command := Command + '(';
          Command := Command + '  DOCTO_CM_DET_ID_MSP,';
          Command := Command + '  DOCTO_CM_ID_MSP,';
          Command := Command + '  NOMBRE,';
          Command := Command + '  UNIDADES,';
          Command := Command + '  PRECIO_UNITARIO,';
          Command := Command + '  PCTJE_DSCTO,';
          Command := Command + '  PRECIO_TOTAL_NETO,';
          Command := Command + '  NOTAS,';
          Command := Command + '  POSICION,';
          Command := Command + '  EMP_FK';
          Command := Command + ')';
          Command := Command + 'VALUES';
          Command := Command + '(';
          Command := Command + '  ' + docto_det + ',';
          Command := Command + '  ' + docto + ',';
          Command := Command + '  ' + QuotedStr(nombre) + ',';
          Command := Command + '  ' + unidades + ',';
          Command := Command + '  ' + QuotedStr(precio_unitario) + ',';
          Command := Command + '  ' + QuotedStr(pctje_dscto) + ',';
          Command := Command + '  ' + QuotedStr(precio_total) + ',';
          Command := Command + '  ' + QuotedStr(notas) + ',';
          Command := Command + '  ' + posicion + ',';
          Command := Command + '  ' + empresa_id;
          Command := Command + ')';

          D.MySQL_Command.CommandText := Command;
          D.MySQL_Command.Execute;
        except
          on E : Exception do
            begin
              Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los detalles de las recepciones|' + folio + '|' + empresa_id + '|' + Command);
            end;
        end;

        D.IBQueryDetalle.Next;
      end;

    D.IBQueryDetalle.Active := False;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los detalles de las recepciones');
      end;
  end;
end;
{$ENDREGION}

{$REGION 'ACTUALIZA_RECEPCIONES - ACTUALIZA LAS RECEPCIONES DE UNA EMPRESA ESPECIFICADA'}
Function ACTUALIZA_RECEPCIONES():Boolean;
  var
    Fmt :TFormatSettings;
    docto_id_old, recep_id, docto, folio, fecha, clve_prov, proveedor, moneda, nombre, simbol, imporneto, impuestos, retencion, modif, dias, almacen, uso_cfdi, Empresa_ID, Empresa_NOMBRE :string;
    ConnectionString, Command :string;
begin
  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat := 'hh:nn:ss';
  Fmt.TimeSeparator := ':';

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Recepciones')) then
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

        D.JvCsvDataSet_Recepcion.Close;
        D.JvCsvDataSet_Recepcion.FileName := ExtractFilePath(ParamStr(0)) + '/Update/Recepciones';
        D.JvCsvDataSet_Recepcion.Open;
        D.JvCsvDataSet_Recepcion.First;
        while not D.JvCsvDataSet_Recepcion.Eof do
          begin
            Application.ProcessMessages;
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Buscando recepción ' + D.JvCsvDataSet_Recepcion.FieldByName('FOLIO').AsString, '', '');

            try
              // CONEXIÓN MICROSIP (SI ES QUE CAMBIA DE EMPRESA ESTO ES PARA BUSCAR EL DETALLE DE LAS RECEPCIONES)
              if (D.Conexion_Microsip.DatabaseName <> (D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Recepcion.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB')) then
                begin
                  D.Conexion_Microsip.Connected := False;
                  D.Conexion_Microsip.DatabaseName := D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Recepcion.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB';
                  D.Conexion_Microsip.Connected := True;
                  D.Transaction_Microsip.Active := True;
                end;

              docto := D.JvCsvDataSet_Recepcion.FieldByName('DOCTO_CM_ID').AsString;
              folio := D.JvCsvDataSet_Recepcion.FieldByName('FOLIO').AsString;
              fecha := FormatDateTime('YYYY-MM-DD HH:NN:SS', StrToDateTime(D.JvCsvDataSet_Recepcion.FieldByName('FECHA').AsString, Fmt));
              clve_prov := D.JvCsvDataSet_Recepcion.FieldByName('CLAVE_PROV').AsString;
              proveedor := D.JvCsvDataSet_Recepcion.FieldByName('PROVEEDOR_ID').AsString;
              moneda := D.JvCsvDataSet_Recepcion.FieldByName('MONEDA_ID').AsString;
              nombre := D.JvCsvDataSet_Recepcion.FieldByName('NOMBRE').AsString;
              simbol := D.JvCsvDataSet_Recepcion.FieldByName('CLAVE_FISCAL').AsString;
              imporneto := StringReplace(FormatFloat('#.000000', D.JvCsvDataSet_Recepcion.FieldByName('IMPORTE_NETO').AsFloat), ',', '.', [rfReplaceAll]);
              impuestos := StringReplace(FormatFloat('#.000000', D.JvCsvDataSet_Recepcion.FieldByName('TOTAL_IMPUESTOS').AsFloat), ',', '.', [rfReplaceAll]);
              retencion := StringReplace(FormatFloat('#.000000', D.JvCsvDataSet_Recepcion.FieldByName('TOTAL_RETENCIONES').AsFloat), ',', '.', [rfReplaceAll]);
              modif := FormatDateTime('YYYY-MM-DD HH:NN:SS', StrToDateTime(D.JvCsvDataSet_Recepcion.FieldByName('FECHA_HORA_ULT_MODIF').AsString, Fmt));
              dias := GET_PLAZO(docto);
              almacen := D.JvCsvDataSet_Recepcion.FieldByName('ALMACEN_ID').AsString;
              uso_cfdi := D.JvCsvDataSet_Recepcion.FieldByName('USO_CFDI').AsString;
              Empresa_ID := D.JvCsvDataSet_Recepcion.FieldByName('EMPRESA_ID').AsString;
              Empresa_NOMBRE := D.JvCsvDataSet_Recepcion.FieldByName('EMPRESA_NOMBRE').AsString;

              // CONEXIÓN MYSQL
              D.ADOQueryActual.Active := False;
              D.ADOQueryActual.SQL.Clear;
              D.ADOQueryActual.SQL.Add('SELECT * FROM RECEPCIONES');
              D.ADOQueryActual.SQL.Add(' WHERE FOLIO = ''' + folio + '''');
              D.ADOQueryActual.SQL.Add('   AND EMP_FK = ' + Empresa_ID);
              D.ADOQueryActual.Active := True;
              D.ADOQueryActual.First;

              docto_id_old := D.ADOQueryActual.FieldByName('DOCTO_CM_ID').AsString;
              recep_id := D.ADOQueryActual.FieldByName('RECEP_ID').AsString;

              // SI LA RECEPCIÓN ESTA PENDIENTE
              if (D.JvCsvDataSet_Recepcion.FieldByName('ESTATUS').AsString = 'P') then
                begin
                  if (D.ADOQueryActual.RecordCount = 0) then
                    begin
                      Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Registrando recepción ' + D.JvCsvDataSet_Recepcion.FieldByName('FOLIO').AsString, '', '');

                      {$REGION 'INSERTA LA NUEVA RECEPCIÓN'}
                      Command := 'INSERT INTO RECEPCIONES';
                      Command := Command + '(';
                      Command := Command + '  DOCTO_CM_ID,';
                      Command := Command + '  FOLIO,';
                      Command := Command + '  FECHA,';
                      Command := Command + '  CLAVE_PROV,';
                      Command := Command + '  PROVEEDOR_ID,';
                      Command := Command + '  MONEDA_ID,';
                      Command := Command + '  MONEDA_NOMBRE,';
                      Command := Command + '  MONEDA_SIMBOLO,';
                      Command := Command + '  IMPORTE_NETO,';
                      Command := Command + '  TOTAL_IMPUESTOS,';
                      Command := Command + '  TOTAL_RETENCIONES,';
                      Command := Command + '  FECHA_HORA_ULT_MODIF,';
                      Command := Command + '  EMP_FK,';
                      Command := Command + '  ESTATUS,';
                      Command := Command + '  DIAS_PLAZO,';
                      Command := Command + '  ALMACEN_FK_MSP,';
                      Command := Command + '  USO_CFDI';
                      Command := Command + ')';
                      Command := Command + 'VALUES';
                      Command := Command + '(';
                      Command := Command + '  ' + docto + ',';
                      Command := Command + '  ' + QuotedStr(folio) + ',';
                      Command := Command + '  ' + QuotedStr(fecha) + ',';
                      Command := Command + '  ' + QuotedStr(clve_prov) + ',';
                      Command := Command + '  ' + proveedor + ',';
                      Command := Command + '  ' + moneda + ',';
                      Command := Command + '  ' + QuotedStr(nombre) + ',';
                      Command := Command + '  ' + QuotedStr(simbol) + ',';
                      Command := Command + '  ' + QuotedStr(imporneto) + ',';
                      Command := Command + '  ' + QuotedStr(impuestos) + ',';
                      Command := Command + '  ' + QuotedStr(retencion) + ',';
                      Command := Command + '  ' + QuotedStr(modif) + ',';
                      Command := Command + '  ' + Empresa_ID + ',';
                      Command := Command + '  ' + QuotedStr('P') + ',';
                      Command := Command + '  ' + dias + ',';
                      Command := Command + '  ' + almacen + ',';
                      Command := Command + '  ' + QuotedStr(uso_cfdi);
                      Command := Command + ')';

                      D.MySQL_Command.CommandText := Command;
                      D.MySQL_Command.Execute;
                      {$ENDREGION}

                      ACTUALIZA_RECEPCIONES_DET('', folio, Empresa_ID);
                    end
                  else
                    begin
                      Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando recepción ' + D.JvCsvDataSet_Recepcion.FieldByName('FOLIO').AsString, '', '');

                      {$REGION 'ACTUALIZA LA RECEPCIÓN EN EL PORTAL'}
                      Command := 'UPDATE RECEPCIONES SET ';
                      Command := Command + '      DOCTO_CM_ID = ' + docto + ',';
                      Command := Command + '      FECHA = ' + QuotedStr(fecha) + ',';
                      Command := Command + '      CLAVE_PROV = ' + QuotedStr(clve_prov) + ',';
                      Command := Command + '      PROVEEDOR_ID = ' + proveedor + ',';
                      Command := Command + '      MONEDA_ID = ' + moneda + ',';
                      Command := Command + '      MONEDA_NOMBRE = ' + QuotedStr(nombre) + ',';
                      Command := Command + '      MONEDA_SIMBOLO = ' + QuotedStr(simbol) + ',';
                      Command := Command + '      IMPORTE_NETO = ' + QuotedStr(imporneto) + ',';
                      Command := Command + '      TOTAL_IMPUESTOS = ' + QuotedStr(impuestos) + ',';
                      Command := Command + '      TOTAL_RETENCIONES = ' + QuotedStr(retencion) + ',';
                      Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(modif) + ',';
                      Command := Command + '      USO_CFDI = ' + QuotedStr(uso_cfdi) + ' ';
                      Command := Command + 'WHERE FOLIO = ' + QuotedStr(folio) + ' ';
                      Command := Command + '  AND EMP_FK = ' + Empresa_ID;

                      D.MySQL_Command.CommandText := Command;
                      D.MySQL_Command.Execute;
                      {$ENDREGION}

                      {$REGION 'SI YA TIENE FACTURA LE INDICA EL NUEVO ID DE LA RECEPCIÓN'}
                      Command := 'UPDATE FACTURA_PROVEEDOR SET ';
                      Command := Command + '      RECEPCION_ID = ' + docto + ' ';
                      Command := Command + 'WHERE RECEP_ID = ' + recep_id;

                      D.MySQL_Command.CommandText := Command;
                      D.MySQL_Command.Execute;
                      {$ENDREGION}

                      ACTUALIZA_RECEPCIONES_DET(docto_id_old, folio, Empresa_ID);
                    end;
                end;

              // SI LA RECEPCIÓN YA ESTA FACTURADA
              if (D.JvCsvDataSet_Recepcion.FieldByName('ESTATUS').AsString = 'F') then
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando recepción ' + D.JvCsvDataSet_Recepcion.FieldByName('FOLIO').AsString, '', '');

                  {$REGION 'MARCA COMO FACTURADA LA RECEPCIÓN EN EL PORTAL'}
                  Command := 'UPDATE RECEPCIONES SET ';
                  Command := Command + '      DOCTO_CM_ID = ' + docto + ',';
                  Command := Command + '      ESTATUS = ' + QuotedStr('F') + ',';
                  Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(modif) + ' ';
                  Command := Command + 'WHERE FOLIO = ' + QuotedStr(folio) + ' ';
                  Command := Command + '  AND EMP_FK = ' + Empresa_ID;

                  D.MySQL_Command.CommandText := Command;
                  D.MySQL_Command.Execute;
                  {$ENDREGION}
                end;

              // SI LA RECEPCIÓN ESTA CANCELADA
              if (D.JvCsvDataSet_Recepcion.FieldByName('ESTATUS').AsString = 'C') then
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Cancelando recepción ' + D.JvCsvDataSet_Recepcion.FieldByName('FOLIO').AsString, '', '');

                  {$REGION 'MARCA COMO CANCELADA LA RECEPCIÓN EN EL PORTAL'}
                  Command := 'UPDATE RECEPCIONES SET ';
                  Command := Command + '      DOCTO_CM_ID = ' + docto + ',';
                  Command := Command + '      ESTATUS = ' + QuotedStr('C') + ',';
                  Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(modif) + ' ';
                  Command := Command + 'WHERE FOLIO = ' + QuotedStr(folio) + ' ';
                  Command := Command + '  AND EMP_FK = ' + Empresa_ID;

                  D.MySQL_Command.CommandText := Command;
                  D.MySQL_Command.Execute;
                  {$ENDREGION}
                end;
            except
              on E : Exception do
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar las recepciones');
                end;
            end;

            Inc(D.Position);
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

            D.JvCsvDataSet_Recepcion.Next;
          end;

        // NOS DESCONECTAMOS
        D.Transaction_Microsip.Active := False;
        D.Conexion_Microsip.Connected := False;

        Result := True;
        D.JvCsvDataSet_Recepcion.Close;
      except
        on E : Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar las recepciones');
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Recepciones'));
    end
  else
    begin
      Result := True;
    end;
end;
{$ENDREGION}

end.
