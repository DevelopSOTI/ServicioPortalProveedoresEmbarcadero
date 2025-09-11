unit Func_Creditos;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf;

  Function ACTUALIZA_CREDITOS():Boolean;

implementation

uses
  Data, Func, Form;

{$REGION 'VALIDA_COMPLEMENTO - VALIDA SI EL COBRO REQUIERE COMPLEMENTO DE PAGO O NO'}
Function VALIDA_COMPLEMENTO(folio, concepto, empresa_id :string; Fmt :TFormatSettings):Boolean;
  var
    impte_docto_cp_id, docto_cp_id, docto_cp_acr_id, sistema_origen, xml, uuid, fecha :string;
begin
  Result := False;

  try
    D.IBQueryDetalle.Active := False;
    D.IBQueryDetalle.SQL.Clear;
    D.IBQueryDetalle.SQL.Add('SELECT');
    D.IBQueryDetalle.SQL.Add('       id.impte_docto_cp_id,');
    D.IBQueryDetalle.SQL.Add('       id.docto_cp_id,');
    D.IBQueryDetalle.SQL.Add('       id.docto_cp_acr_id,');
    D.IBQueryDetalle.SQL.Add('       dc.sistema_origen');
    D.IBQueryDetalle.SQL.Add('  FROM doctos_cp de');
    D.IBQueryDetalle.SQL.Add('  JOIN importes_doctos_cp id ON(de.docto_cp_id = id.docto_cp_id)');
    D.IBQueryDetalle.SQL.Add('  JOIN doctos_cp dc ON(id.docto_cp_acr_id = dc.docto_cp_id)');
    D.IBQueryDetalle.SQL.Add(' WHERE de.folio = ''' + folio + '''');
    D.IBQueryDetalle.SQL.Add('   AND de.concepto_cp_id = ' + concepto);
    D.IBQueryDetalle.Active := True;
    D.IBQueryDetalle.First;

    while not D.IBQueryDetalle.Eof do
      begin
        Application.ProcessMessages;

        try
          impte_docto_cp_id := D.IBQueryDetalle.FieldByName('IMPTE_DOCTO_CP_ID').AsString;
          docto_cp_id := D.IBQueryDetalle.FieldByName('DOCTO_CP_ID').AsString;
          docto_cp_acr_id := D.IBQueryDetalle.FieldByName('DOCTO_CP_ACR_ID').AsString;
          sistema_origen := D.IBQueryDetalle.FieldByName('SISTEMA_ORIGEN').AsString;

          D.IBQueryXML.Active := False;
          D.IBQueryXML.SQL.Clear;

          if (sistema_origen = 'CM') then
            begin
              D.IBQueryXML.SQL.Add('SELECT');
              D.IBQueryXML.SQL.Add('       cr.xml,');
              D.IBQueryXML.SQL.Add('       rc.uuid,');
              D.IBQueryXML.SQL.Add('       rc.fecha');
              D.IBQueryXML.SQL.Add('  FROM doctos_entre_sis ds');
              D.IBQueryXML.SQL.Add('  JOIN cfd_recibidos cr ON(ds.docto_fte_id = cr.docto_id)');
              D.IBQueryXML.SQL.Add('  JOIN repositorio_cfdi rc ON(cr.cfdi_id = rc.cfdi_id)');
              D.IBQueryXML.SQL.Add(' WHERE ds.docto_dest_id = ' + docto_cp_acr_id);
              D.IBQueryXML.SQL.Add('   AND cr.clave_sistema = ''' + sistema_origen + '''');
            end
          else
            begin
              D.IBQueryXML.SQL.Add('SELECT');
              D.IBQueryXML.SQL.Add('       cr.xml,');
              D.IBQueryXML.SQL.Add('       rc.uuid,');
              D.IBQueryXML.SQL.Add('       rc.fecha');
              D.IBQueryXML.SQL.Add('  FROM cfd_recibidos cr');
              D.IBQueryXML.SQL.Add('  JOIN repositorio_cfdi rc ON(cr.cfdi_id = rc.cfdi_id)');
              D.IBQueryXML.SQL.Add(' WHERE cr.docto_id = ' + docto_cp_acr_id);
              D.IBQueryXML.SQL.Add('   AND cr.clave_sistema = ''' + sistema_origen + '''');
            end;

          D.IBQueryXML.Active := True;
          D.IBQueryXML.First;
          while not D.IBQueryXML.Eof do
            begin
              xml := D.IBQueryXML.FieldByName('XML').AsString;
              uuid := D.IBQueryXML.FieldByName('UUID').AsString;
              // fecha := FormatDateTime('YYYY-MM-DD', StrToDateTime(D.IBQueryXML.FieldByName('FECHA').AsString, Fmt));
              fecha := FormatDateTime('YYYY-MM-DD', D.IBQueryXML.FieldByName('FECHA').AsDateTime);

              xml := UTF8ToString(xml);
              D.XML_FILE.LoadFromXML(xml);
              D.XML_FILE.Version := '1.0';

              if (D.XML_FILE.ChildNodes.FindNode('cfdi:Comprobante').Attributes['Version'] = '3.3') then
                begin
                  if (D.XML_FILE.ChildNodes.FindNode('cfdi:Comprobante').Attributes['MetodoPago'] = 'PPD') then
                    begin
                      // CON QUE ENCUENTRE UN PPD EN EL COBRO DEVUELVE VERDADERO
                      Result := True;
                    end;
                end;

              if (D.XML_FILE.ChildNodes.FindNode('cfdi:Comprobante').Attributes['Version'] = '4.0') then
                begin
                  if (D.XML_FILE.ChildNodes.FindNode('cfdi:Comprobante').Attributes['MetodoPago'] = 'PPD') then
                    begin
                      // CON QUE ENCUENTRE UN PPD EN EL COBRO DEVUELVE VERDADERO
                      Result := True;
                    end;
                end;

              D.IBQueryXML.Next;
            end;
        except
          on E : Exception do
            begin
              Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los detalles de los creditos');
              Result := False;
            end;
        end;

        D.IBQueryDetalle.Next;
      end;

    D.IBQueryDetalle.Active := False;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los detalles de los creditos');
        Result := False;
      end;
  end;
end;
{$ENDREGION}

{$REGION 'ACTUALIZA_CREDITOS_DET - ACTUALIZA EL DETALLE DEL COBRO EN PROCESO'}
Function ACTUALIZA_CREDITOS_DET(docto_id_old, credito_id, folio, concepto, empresa_id :string; Fmt :TFormatSettings):Boolean;
  var
    impte_docto_cp_id, docto_cp_id, docto_cp_acr_id, importe, impuesto, iva_retenido, isr_retenido, folio_acr, descripcion, sistema_origen, xml, uuid, fecha :string;
    Command :string;
begin
  Result := False;

  if (docto_id_old <> '') then
    begin
      try
        Command := 'DELETE FROM CREDITOS_DET ';
        Command := Command + 'WHERE DOCTO_CP_ID = ' + docto_id_old + ' ';
        Command := Command + '  AND EMP_FK = ' + empresa_id;

        D.MySQL_Command.CommandText := Command;
        D.MySQL_Command.Execute;
      except
        on E : Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los detalles de los creditos');
          end;
      end;
    end;

  try
    D.IBQueryDetalle.Active := False;
    D.IBQueryDetalle.SQL.Clear;
    D.IBQueryDetalle.SQL.Add('SELECT');
    D.IBQueryDetalle.SQL.Add('       id.impte_docto_cp_id,');
    D.IBQueryDetalle.SQL.Add('       id.docto_cp_id,');
    D.IBQueryDetalle.SQL.Add('       id.docto_cp_acr_id,');
    D.IBQueryDetalle.SQL.Add('       id.importe,');
    D.IBQueryDetalle.SQL.Add('       id.impuesto,');
    D.IBQueryDetalle.SQL.Add('       id.iva_retenido,');
    D.IBQueryDetalle.SQL.Add('       id.isr_retenido,');
    D.IBQueryDetalle.SQL.Add('       dc.folio,');
    D.IBQueryDetalle.SQL.Add('       dc.descripcion,');
    D.IBQueryDetalle.SQL.Add('       dc.sistema_origen');
    D.IBQueryDetalle.SQL.Add('  FROM doctos_cp de');
    D.IBQueryDetalle.SQL.Add('  JOIN importes_doctos_cp id ON(de.docto_cp_id = id.docto_cp_id)');
    D.IBQueryDetalle.SQL.Add('  JOIN doctos_cp dc ON(id.docto_cp_acr_id = dc.docto_cp_id)');
    D.IBQueryDetalle.SQL.Add(' WHERE de.folio = ''' + folio + '''');
    D.IBQueryDetalle.SQL.Add('   AND de.concepto_cp_id = ' + concepto);
    D.IBQueryDetalle.Active := True;
    D.IBQueryDetalle.First;

    while not D.IBQueryDetalle.Eof do
      begin
        Application.ProcessMessages;

        try
          impte_docto_cp_id := D.IBQueryDetalle.FieldByName('IMPTE_DOCTO_CP_ID').AsString;
          docto_cp_id := D.IBQueryDetalle.FieldByName('DOCTO_CP_ID').AsString;
          docto_cp_acr_id := D.IBQueryDetalle.FieldByName('DOCTO_CP_ACR_ID').AsString;
          importe := StringReplace(FormatFloat('#.000000', D.IBQueryDetalle.FieldByName('IMPORTE').AsFloat), ',', '.', [rfReplaceAll]);
          impuesto := StringReplace(FormatFloat('#.000000', D.IBQueryDetalle.FieldByName('IMPUESTO').AsFloat), ',', '.', [rfReplaceAll]);
          iva_retenido := StringReplace(FormatFloat('#.000000', D.IBQueryDetalle.FieldByName('IVA_RETENIDO').AsFloat), ',', '.', [rfReplaceAll]);
          isr_retenido := StringReplace(FormatFloat('#.000000', D.IBQueryDetalle.FieldByName('ISR_RETENIDO').AsFloat), ',', '.', [rfReplaceAll]);
          folio_acr := D.IBQueryDetalle.FieldByName('FOLIO').AsString;
          descripcion := D.IBQueryDetalle.FieldByName('DESCRIPCION').AsString.Replace('''', '');
          sistema_origen := D.IBQueryDetalle.FieldByName('SISTEMA_ORIGEN').AsString;

          D.IBQueryXML.Active := False;
          D.IBQueryXML.SQL.Clear;

          if (sistema_origen = 'CM') then
            begin
              D.IBQueryXML.SQL.Add('SELECT');
              D.IBQueryXML.SQL.Add('       cr.xml,');
              D.IBQueryXML.SQL.Add('       rc.uuid,');
              D.IBQueryXML.SQL.Add('       rc.fecha');
              D.IBQueryXML.SQL.Add('  FROM doctos_entre_sis ds');
              D.IBQueryXML.SQL.Add('  JOIN cfd_recibidos cr ON(ds.docto_fte_id = cr.docto_id)');
              D.IBQueryXML.SQL.Add('  JOIN repositorio_cfdi rc ON(cr.cfdi_id = rc.cfdi_id)');
              D.IBQueryXML.SQL.Add(' WHERE ds.docto_dest_id = ' + docto_cp_acr_id);
              D.IBQueryXML.SQL.Add('   AND cr.clave_sistema = ''' + sistema_origen + '''');
            end
          else
            begin
              D.IBQueryXML.SQL.Add('SELECT');
              D.IBQueryXML.SQL.Add('       cr.xml,');
              D.IBQueryXML.SQL.Add('       rc.uuid,');
              D.IBQueryXML.SQL.Add('       rc.fecha');
              D.IBQueryXML.SQL.Add('  FROM cfd_recibidos cr');
              D.IBQueryXML.SQL.Add('  JOIN repositorio_cfdi rc ON(cr.cfdi_id = rc.cfdi_id)');
              D.IBQueryXML.SQL.Add(' WHERE cr.docto_id = ' + docto_cp_acr_id);
              D.IBQueryXML.SQL.Add('   AND cr.clave_sistema = ''' + sistema_origen + '''');
            end;

          D.IBQueryXML.Active := True;
          D.IBQueryXML.First;
          while not D.IBQueryXML.Eof do
            begin
              xml := D.IBQueryXML.FieldByName('XML').AsString;
              uuid := D.IBQueryXML.FieldByName('UUID').AsString;
              // fecha := FormatDateTime('YYYY-MM-DD', StrToDateTime(D.IBQueryXML.FieldByName('FECHA').AsString, Fmt));
              // fecha := FormatDateTime('YYYY-MM-DD', StrToDateTime(D.IBQueryXML.FieldByName('FECHA').AsString, Fmt));
              fecha := FormatDateTime('YYYY-MM-DD', D.IBQueryXML.FieldByName('FECHA').AsDateTime);

              xml := UTF8ToString(xml);
              D.XML_FILE.LoadFromXML(xml);
              D.XML_FILE.Version := '1.0';

              if (D.XML_FILE.ChildNodes.FindNode('cfdi:Comprobante').Attributes['Version'] = '3.3') or (D.XML_FILE.ChildNodes.FindNode('cfdi:Comprobante').Attributes['Version'] = '4.0') then
                begin
                  if (D.XML_FILE.ChildNodes.FindNode('cfdi:Comprobante').Attributes['MetodoPago'] = 'PPD') then
                    begin
                      Command := 'INSERT INTO CREDITOS_DET';
                      Command := Command + '(';
                      Command := Command + '  IMPTE_DOCTO_CP_ID,';
                      Command := Command + '  CREDITO_FK,';
                      Command := Command + '  DOCTO_CP_ID,';
                      Command := Command + '  DOCTO_CP_ACR_ID,';
                      Command := Command + '  IMPORTE,';
                      Command := Command + '  IMPUESTO,';
                      Command := Command + '  IVA_RETENIDO,';
                      Command := Command + '  ISR_RETENIDO,';
                      Command := Command + '  FOLIO_ACR,';
                      Command := Command + '  DESCRIPCION,';
                      Command := Command + '  UUID,';
                      Command := Command + '  FECHA,';
                      Command := Command + '  EMP_FK';
                      Command := Command + ')';
                      Command := Command + 'VALUES';
                      Command := Command + '(';
                      Command := Command + '  ' + impte_docto_cp_id + ',';
                      Command := Command + '  ' + credito_id + ',';
                      Command := Command + '  ' + docto_cp_id + ',';
                      Command := Command + '  ' + docto_cp_acr_id + ',';
                      Command := Command + '  ' + QuotedStr(importe) + ',';
                      Command := Command + '  ' + QuotedStr(impuesto) + ',';
                      Command := Command + '  ' + QuotedStr(iva_retenido) + ',';
                      Command := Command + '  ' + QuotedStr(isr_retenido) + ',';
                      Command := Command + '  ' + QuotedStr(folio_acr) + ',';
                      // Command := Command + '  ' + QuotedStr(descripcion) + ',';
                      Command := Command + '  :DESCRIPCION,';
                      Command := Command + '  ' + QuotedStr(uuid) + ',';
                      Command := Command + '  ' + QuotedStr(fecha) + ',';
                      Command := Command + '  ' + empresa_id;
                      Command := Command + ')';

                      D.MySQL_Command.CommandText := Command;
                      D.MySQL_Command.Parameters.ParamByName('DESCRIPCION').Value := descripcion;
                      D.MySQL_Command.Execute;

                      // CON QUE ENCUENTRE UN PPD EN EL COBRO DEVUELVE VERDADERO
                      Result := True;
                    end;
                end;

              D.IBQueryXML.Next;
            end;
        except
          on E : Exception do
            begin
              Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar los detalles del credito ' + folio);
              Result := False;
            end;
        end;

        D.IBQueryDetalle.Next;
      end;

    D.IBQueryDetalle.Active := False;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al revisar los detalles del credito ' + folio);
        Result := False;
      end;
  end;
end;
{$ENDREGION}

Function ACTUALIZA_CREDITOS():Boolean;
  var
    Resultado :Boolean;
    Fmt :TFormatSettings;
    docto_id_old, credito_id, docto, concepto, nombre, descripcion, folio, fecha, clve_prov, proveedor, modif, Empresa_ID, Empresa_NOMBRE :string;
    Command :string;
begin
  Resultado := True;

  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat := 'hh:nn:ss';
  Fmt.TimeSeparator := ':';

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Creditos')) then
    begin
      try
        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Credito.Close;
        D.JvCsvDataSet_Credito.FileName := ExtractFilePath(ParamStr(0)) + '/Update/Creditos';
        D.JvCsvDataSet_Credito.Open;
        D.JvCsvDataSet_Credito.First;
        while not D.JvCsvDataSet_Credito.Eof do
          begin
            Application.ProcessMessages;
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Buscando credito ' + D.JvCsvDataSet_Credito.FieldByName('FOLIO').AsString, '', '');

            try
              // CONEXIÓN MICROSIP (SI ES QUE CAMBIA DE EMPRESA)
              if (D.Conexion_Microsip.DatabaseName <> (D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Credito.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB')) then
                begin
                  D.Conexion_Microsip.Connected := False;
                  D.Conexion_Microsip.DatabaseName := D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Credito.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB';
                  D.Conexion_Microsip.Connected := True;
                  D.Transaction_Microsip.Active := True;
                end;

              docto := D.JvCsvDataSet_Credito.FieldByName('DOCTO_CP_ID').AsString;
              concepto := D.JvCsvDataSet_Credito.FieldByName('CONCEPTO_CP_ID').AsString;
              nombre := D.JvCsvDataSet_Credito.FieldByName('CONCEPTO_CP').AsString;
              folio := D.JvCsvDataSet_Credito.FieldByName('FOLIO').AsString;
              fecha := FormatDateTime('YYYY-MM-DD HH:NN:SS', StrToDateTime(D.JvCsvDataSet_Credito.FieldByName('FECHA').AsString, Fmt));
              clve_prov := D.JvCsvDataSet_Credito.FieldByName('CLAVE_PROV').AsString;
              proveedor := D.JvCsvDataSet_Credito.FieldByName('PROVEEDOR_ID').AsString;
              modif := FormatDateTime('YYYY-MM-DD HH:NN:SS', StrToDateTime(D.JvCsvDataSet_Credito.FieldByName('FECHA_HORA_ULT_MODIF').AsString, Fmt));
              descripcion := D.JvCsvDataSet_Credito.FieldByName('DESCRIPCION').AsString;
              Empresa_ID := D.JvCsvDataSet_Credito.FieldByName('EMPRESA_ID').AsString;
              Empresa_NOMBRE := D.JvCsvDataSet_Credito.FieldByName('EMPRESA_NOMBRE').AsString;

              // BUSCAMOS EL CREDITO EN EL PORTAL
              D.ADOQueryActual.Active := False;
              D.ADOQueryActual.SQL.Clear;
              D.ADOQueryActual.SQL.Add('SELECT * FROM CREDITOS');
              D.ADOQueryActual.SQL.Add(' WHERE FOLIO = ''' + folio + '''');
              D.ADOQueryActual.SQL.Add('   AND CONCEPTO_CP_ID = ' + concepto);
              D.ADOQueryActual.SQL.Add('   AND PROVEEDOR_ID = ' + proveedor);
              D.ADOQueryActual.SQL.Add('   AND EMP_FK = ' + Empresa_ID);
              D.ADOQueryActual.Active := True;
              D.ADOQueryActual.First;

              docto_id_old := D.ADOQueryActual.FieldByName('DOCTO_CP_ID').AsString;
              credito_id := D.ADOQueryActual.FieldByName('CREDITO_ID').AsString;

              if (D.JvCsvDataSet_Credito.FieldByName('TIENE_CFD').AsString = 'N') and (D.JvCsvDataSet_Credito.FieldByName('CANCELADO').AsString = 'N') then
                begin
                  if (D.ADOQueryActual.RecordCount = 0) then
                    begin
                      Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Registrando credito ' + D.JvCsvDataSet_Credito.FieldByName('FOLIO').AsString, '', '');

                      if (VALIDA_COMPLEMENTO(folio, concepto, Empresa_ID, Fmt) = True) then
                        begin
                          try
                            {$REGION 'REGISTRA EL ENCABEZADO DEL CREDITO EN EL PORTAL'}
                            Command := 'INSERT INTO CREDITOS';
                            Command := Command + '(';
                            Command := Command + '  DOCTO_CP_ID,';
                            Command := Command + '  CONCEPTO_CP_ID,';
                            Command := Command + '  FOLIO,';
                            Command := Command + '  FECHA,';
                            Command := Command + '  CLAVE_PROV,';
                            Command := Command + '  PROVEEDOR_ID,';
                            Command := Command + '  FECHA_HORA_ULT_MODIF,';
                            Command := Command + '  EMP_FK,';
                            Command := Command + '  APLICADO,';
                            Command := Command + '  CANCELADO,';
                            Command := Command + '  ESTATUS,';
                            Command := Command + '  CONCEPTO_CP,';
                            Command := Command + '  DESCRIPCION';
                            Command := Command + ')';
                            Command := Command + 'VALUES';
                            Command := Command + '(';
                            Command := Command + '  ' + docto + ',';
                            Command := Command + '  ' + concepto + ',';
                            Command := Command + '  ''' + folio + ''',';
                            Command := Command + '  ''' + fecha + ''',';
                            Command := Command + '  ''' + clve_prov + ''',';
                            Command := Command + '  ' + proveedor + ',';
                            Command := Command + '  ''' + modif + ''',';
                            Command := Command + '  ' + Empresa_ID + ',';
                            Command := Command + '  ''S'',';
                            Command := Command + '  ''N'', ';
                            Command := Command + '  ''P'',';
                            Command := Command + '  ''' + nombre + ''',';
                            Command := Command + '  ''' + descripcion + '''';
                            Command := Command + ')';

                            D.MySQL_Command.CommandText := Command;
                            D.MySQL_Command.Execute;

                            D.ADOQueryActual.Active := False;
                            D.ADOQueryActual.SQL.Clear;
                            D.ADOQueryActual.SQL.Add('SELECT * FROM CREDITOS');
                            D.ADOQueryActual.SQL.Add(' WHERE FOLIO = ''' + folio + '''');
                            D.ADOQueryActual.SQL.Add('   AND CONCEPTO_CP_ID = ' + concepto);
                            D.ADOQueryActual.SQL.Add('   AND PROVEEDOR_ID = ' + proveedor);
                            D.ADOQueryActual.SQL.Add('   AND EMP_FK = ' + empresa_id);
                            D.ADOQueryActual.Active := True;
                            D.ADOQueryActual.First;

                            credito_id := D.ADOQueryActual.FieldByName('CREDITO_ID').AsString;
                            {$ENDREGION}

                            ACTUALIZA_CREDITOS_DET('', credito_id, folio, concepto, Empresa_ID, Fmt);
                          except
                            on E : Exception do
                              begin
                                Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al registrar el credito ' + folio);
                                Resultado := False;
                              end;
                          end;
                        end;
                    end
                  else
                    begin
                      Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Actualizando credito ' + D.JvCsvDataSet_Credito.FieldByName('FOLIO').AsString, '', '');

                      if (VALIDA_COMPLEMENTO(folio, concepto, Empresa_ID, Fmt) = True) then
                        begin
                          try
                            {$REGION 'ACTUALIZA EL ENCABEZADO DEL CREDITO EN EL PORTAL'}
                            Command := 'UPDATE CREDITOS SET ';
                            Command := Command + '      DOCTO_CP_ID = ' + docto + ',';
                            Command := Command + '      FECHA = ' + QuotedStr(fecha) + ',';
                            Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(modif) + ',';
                            Command := Command + '      DESCRIPCION = ' + QuotedStr(descripcion) + ' ';
                            Command := Command + 'WHERE CREDITO_ID = ' + credito_id;

                            D.MySQL_Command.CommandText := Command;
                            D.MySQL_Command.Execute;
                            {$ENDREGION}

                            ACTUALIZA_CREDITOS_DET(docto_id_old, credito_id, folio, concepto, Empresa_ID, Fmt);
                          except
                            on E : Exception do
                              begin
                                Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar el credito ' + folio);
                                Resultado := False;
                              end;
                          end;
                        end;
                    end;
                end;

              // SI EL CREDITO YA TIENE UN CFDI ASOCIADO LO QUITA DEL PORTAL
              if (D.JvCsvDataSet_Credito.FieldByName('TIENE_CFD').AsString = 'S') and (credito_id <> '') then
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Finalizando credito ' + D.JvCsvDataSet_Credito.FieldByName('FOLIO').AsString, '', '');

                  Command := 'UPDATE CREDITOS SET ';
                  Command := Command + '      DOCTO_CP_ID = ' + docto + ', ';
                  Command := Command + '      ESTATUS = ''F'', ';
                  Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(modif) + ' ';
                  Command := Command + 'WHERE CREDITO_ID = ' + credito_id;

                  D.MySQL_Command.CommandText := Command;
                  D.MySQL_Command.Execute;
                end;

              // SI EL CREDITO ESTA CANCELADO
              if (D.JvCsvDataSet_Credito.FieldByName('CANCELADO').AsString = 'S') and (credito_id <> '') then
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Cancelando credito ' + D.JvCsvDataSet_Credito.FieldByName('FOLIO').AsString, '', '');

                  Command := 'UPDATE CREDITOS SET ';
                  Command := Command + '      DOCTO_CP_ID = ' + docto + ', ';
                  Command := Command + '      ESTATUS = ''C'', ';
                  Command := Command + '      CANCELADO = ''S'', ';
                  Command := Command + '      FECHA_HORA_ULT_MODIF = ' + QuotedStr(modif) + ' ';
                  Command := Command + 'WHERE CREDITO_ID = ' + credito_id;

                  D.MySQL_Command.CommandText := Command;
                  D.MySQL_Command.Execute;
                end;
            except
              on E : Exception do
                begin
                  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al actualizar el credito ' + folio);
                  Resultado := False;
                end;
            end;

            Inc(D.Position);
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

            D.JvCsvDataSet_Credito.Next;
          end;

        // NOS DESCONECTAMOS DE MICROSIP
        D.Transaction_Microsip.Active := False;
        D.Conexion_Microsip.Connected := False;

        D.JvCsvDataSet_Credito.Close;
      except
        on E : Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al sincronizar los creditos');
            Resultado := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Creditos'));
    end
  else
    begin
      Resultado := True;
    end;

  Result := Resultado;
end;

end.
