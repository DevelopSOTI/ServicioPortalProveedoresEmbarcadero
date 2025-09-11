unit Func_Calcula;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf;

  // FUNCIONES PARA CALCULAR LOS REGISTROS A REGISTRAR/ACTUALIZAR
  Function CALCULA_REGISTROS(var Total :Integer):Boolean;

implementation

uses
  Data, Func;

{$REGION 'CALCULA_EMPRESAS - CALCULA LOS REGISTROS DE LAS EMPRESAS A ACTUALIZAR'}
Function CALCULA_EMPRESAS(var Total :Integer):Boolean;
  var
    LAST, Line :string;
    List :TStringList;
    Agregar :Boolean;
    Ignoradas :Integer;
begin
  List := TStringList.Create;
  List.Clear;

  try
    // NOS CONECTAMOS AL CONFIG
    D.Conexion_Config.Connected := False;
    D.Conexion_Config.DatabaseName := D.MICRO_SERV + ':' + D.MICRO_ROOT + 'System\Config.FDB';
    D.Conexion_Config.Connected := True;
    D.Transaction_Config.Active := True;

    // OBTENEMOS ULTIMA FECHA DE MODIFICACIÓN CON FORMATO PARA MICROSIP
    LAST := FormatDateTime('dd.mm.yyyy hh:nn:ss', D.LastUpdate_Date);

    // REVISAMOS SI YA EXISTE EL ARCHIVO 'Empresas' SI YA EXISTE LO LEEMOS EN CASO CONTRARIO INICIAMOS UNO NUEVO
    if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Empresas')) then
      begin
        List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Empresas');
      end
    else
      begin
        List.Add('EMPRESA_ID,NOMBRE_CORTO,FECHA_HORA_ULT_MODIF,NOMBRE,RFC');
      end;

    Ignoradas := 0;

    // REALIZAMOS LA CONSULTA DE LAS EMPRESAS A AGREGAR/MODIFICAR
    D.IBQueryConfig.Active := False;
    D.IBQueryConfig.SQL.Clear;
    D.IBQueryConfig.SQL.Add('SELECT * FROM EMPRESAS');
    if (D.LastUpdate_Text <> '') then
      begin
        D.IBQueryConfig.SQL.Add('WHERE FECHA_HORA_CREACION > ''' + LAST + '''');
        D.IBQueryConfig.SQL.Add('   OR FECHA_HORA_ULT_MODIF > ''' + LAST + '''');
      end;
    D.IBQueryConfig.SQL.Add('ORDER BY EMPRESA_ID');
    D.IBQueryConfig.Active := True;
    D.IBQueryConfig.First;
    while not D.IBQueryConfig.Eof do
      begin
        Agregar := True;

        Line := '"' + D.IBQueryConfig.FieldByName('EMPRESA_ID').AsString + '",';
        Line := Line + '"' + D.IBQueryConfig.FieldByName('NOMBRE_CORTO').AsString + '",';
        Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryConfig.FieldByName('FECHA_HORA_ULT_MODIF').AsDateTime) + '",';

        try
          D.Conexion_Microsip.Connected := False;
          D.Conexion_Microsip.DatabaseName := D.MICRO_SERV + ':' + D.MICRO_ROOT + D.IBQueryConfig.FieldByName('NOMBRE_CORTO').AsString + '.FDB';
          D.Conexion_Microsip.Connected := True;
          D.Transaction_Microsip.Active := True;

          // OBTENEMOS EL NOMBRE DE LA EMPRESA
          D.IBQueryMicrosip.Active := False;
          D.IBQueryMicrosip.SQL.Clear;
          D.IBQueryMicrosip.SQL.Add('SELECT VALOR FROM REGISTRY WHERE NOMBRE = ''Nombre''');
          D.IBQueryMicrosip.Active := True;
          Line := Line + '"' + D.IBQueryMicrosip.FieldByName('VALOR').AsString + '",';

          // OBTENEMOS EL RFC DE LA EMPRESA
          D.IBQueryMicrosip.Active := False;
          D.IBQueryMicrosip.SQL.Clear;
          D.IBQueryMicrosip.SQL.Add('SELECT VALOR FROM REGISTRY WHERE NOMBRE = ''Rfc''');
          D.IBQueryMicrosip.Active := True;
          Line := Line + '"' + D.IBQueryMicrosip.FieldByName('VALOR').AsString + '"';
        except
          on E:Exception do
            begin
              Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message);
              Agregar := False;
              Inc(Ignoradas);
            end;
        end;

        D.Transaction_Microsip.Active := False;
        D.Conexion_Microsip.Connected := False;

        if (Agregar) then
          begin
            List.Add(Line);
          end;

        D.IBQueryConfig.Next;
      end;

    List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Empresas');
    Total := D.IBQueryConfig.RecordCount - Ignoradas;

    Result := True;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al revisar las empresas de Microsip.');
        Total := 0;
        Result := False;
      end;
  end;

  D.Transaction_Config.Active := False;
  D.Conexion_Config.Connected := False;

  List.Destroy;
end;
{$ENDREGION}

{$REGION 'CALCULA_REGISTROS_EMPRESAS - CALCULA TODOS LOS REGISTROS A ACTUALIZAR DE TODAS LAS EMPRESAS'}
Function CALCULA_REGISTROS_EMPRESAS(var Total :Integer):Boolean;
  var
    Empresa_ID, Empresa_N, LAST, Line: string;
    // XML :string;
    List :TStringList;
    actualiza_almacenes, actualiza_monedas, actualiza_proveedores, actualiza_recepciones, actualiza_creditos, actualiza_particulares, actualiza_facturas :Boolean;
    descripcion: string;
begin
  List := TStringList.Create;
  Total := 0;

  actualiza_almacenes := False;
  actualiza_monedas := False;
  actualiza_proveedores := False;
  actualiza_recepciones := False;
  actualiza_creditos := True;
  actualiza_particulares := False;
  actualiza_facturas := False;

  try
    D.ADOQueryEmpresas.Active := False;
    D.ADOQueryEmpresas.SQL.Clear;
    D.ADOQueryEmpresas.SQL.Add('SELECT * FROM EMPRESAS_MSP WHERE EMP_ESTATUS = ''Autorizada''');
    D.ADOQueryEmpresas.Active := True;
    D.ADOQueryEmpresas.First;

    while not D.ADOQueryEmpresas.Eof do
      begin
        Empresa_ID := D.ADOQueryEmpresas.FieldByName('EMP_ID_MSP').AsString;
        Empresa_N := D.ADOQueryEmpresas.FieldByName('EMP_NOMBRE').AsString;
        // LAST := FormatDateTime('dd.mm.yyyy hh:nn:ss', D.LastUpdate_Date);

        LAST := '';
        if (D.ADOQueryEmpresas.FieldByName('EMP_ULT_SINC').AsString <> '') then
          begin
            // LAST := FormatDateTime('dd.mm.yyyy hh:nn:ss', D.ADOQueryEmpresas.FieldByName('EMP_ULT_SINC').AsDateTime);
            LAST := FormatDateTime('dd.mm.yyyy hh:nn:ss', IncDay(D.ADOQueryEmpresas.FieldByName('EMP_ULT_SINC').AsDateTime, -1));
          end;

        try
          D.Conexion_Microsip.Connected := False;
          D.Conexion_Microsip.DatabaseName := D.MICRO_SERV + ':' + D.MICRO_ROOT + D.ADOQueryEmpresas.FieldByName('EMP_NOMBRE').AsString + '.FDB';
          D.Conexion_Microsip.Connected := True;
          D.Transaction_Microsip.Active := True;

          if actualiza_almacenes then
            begin
              {$REGION 'REVISA LA CANTIDAD DE ALMACENES A SUBIR O ACTUALIZAR'}

              List.Clear;

              if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Almacenes')) then
                begin
                  List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Almacenes');
                end
              else
                begin
                  List.Add('ALMACEN_ID,NOMBRE,NOMBRE_ABREV,FECHA_HORA_ULT_MODIF,EMPRESA_ID');
                end;

              D.IBQueryMicrosip.Active := False;
              D.IBQueryMicrosip.SQL.Clear;
              D.IBQueryMicrosip.SQL.Add('SELECT ALMACEN_ID, NOMBRE, NOMBRE_ABREV, FECHA_HORA_ULT_MODIF FROM almacenes');

              if (LAST <> '') then
                begin
                  D.IBQueryMicrosip.SQL.Add('WHERE FECHA_HORA_CREACION > ''' + LAST + '''');
                  D.IBQueryMicrosip.SQL.Add('   OR FECHA_HORA_ULT_MODIF > ''' + LAST + '''');
                end;

              D.IBQueryMicrosip.SQL.Add('ORDER BY almacen_id');
              D.IBQueryMicrosip.Active := True;
              D.IBQueryMicrosip.First;
              while not D.IBQueryMicrosip.Eof do
                begin
                  Line := '"' + D.IBQueryMicrosip.FieldByName('ALMACEN_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('NOMBRE').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('NOMBRE_ABREV').AsString + '",';
                  Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryMicrosip.FieldByName('FECHA_HORA_ULT_MODIF').AsDateTime) + '",';
                  Line := Line + '"' + Empresa_ID + '"';

                  List.Add(Line);
                  D.IBQueryMicrosip.Next;
                end;

              List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Almacenes');
              Total := Total + D.IBQueryMicrosip.RecordCount;

              {$ENDREGION}
            end;

          if actualiza_monedas then
            begin
              {$REGION 'REVISA LA CANTIDAD DE MONEDAS A SUBIR O ACTUALIZAR'}

              List.Clear;

              if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Monedas')) then
                begin
                  List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Monedas');
                end
              else
                begin
                  List.Add('MONEDA_ID,NOMBRE,CLAVE_FISCAL,FECHA_HORA_ULT_MODIF,EMPRESA_ID');
                end;

              D.IBQueryMicrosip.Active := False;
              D.IBQueryMicrosip.SQL.Clear;
              D.IBQueryMicrosip.SQL.Add('SELECT MONEDA_ID, NOMBRE, CLAVE_FISCAL, FECHA_HORA_ULT_MODIF FROM monedas');

              if (LAST <> '') then
                begin
                  D.IBQueryMicrosip.SQL.Add('WHERE FECHA_HORA_CREACION > ''' + LAST + '''');
                  D.IBQueryMicrosip.SQL.Add('   OR FECHA_HORA_ULT_MODIF > ''' + LAST + '''');
                end;

              D.IBQueryMicrosip.SQL.Add('ORDER BY moneda_id');
              D.IBQueryMicrosip.Active := True;
              D.IBQueryMicrosip.First;
              while not D.IBQueryMicrosip.Eof do
                begin
                  Line := '"' + D.IBQueryMicrosip.FieldByName('MONEDA_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('NOMBRE').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('CLAVE_FISCAL').AsString + '",';
                  Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryMicrosip.FieldByName('FECHA_HORA_ULT_MODIF').AsDateTime) + '",';
                  Line := Line + '"' + Empresa_ID + '"';

                  List.Add(Line);
                  D.IBQueryMicrosip.Next;
                end;

              List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Monedas');
              Total := Total + D.IBQueryMicrosip.RecordCount;

              {$ENDREGION}
            end;

          if actualiza_proveedores then
            begin
              {$REGION 'REVISA LA CANTIDAD DE PROVEEDORES A SUBIR O ACTUALIZAR'}
              List.Clear;

              if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Proveedores')) then
                begin
                  List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Proveedores');
                end
              else
                begin
                  List.Add('PROVEEDOR_ID,NOMBRE,ESTATUS,CLAVE_PROV,FECHA_HORA_ULT_MODIF,PCTJE_RECHAZO,REFERENCIA,RFC_CURP,PERMITIR_SIN_RECEPCION,EMPRESA_ID');
                end;

              D.IBQueryMicrosip.Active := False;
              D.IBQueryMicrosip.SQL.Clear;
              D.IBQueryMicrosip.SQL.Add('SELECT');
              D.IBQueryMicrosip.SQL.Add('       p.proveedor_id,');
              D.IBQueryMicrosip.SQL.Add('       p.nombre,');
              D.IBQueryMicrosip.SQL.Add('       p.estatus,');
              D.IBQueryMicrosip.SQL.Add('       p.fecha_hora_ult_modif,');
              D.IBQueryMicrosip.SQL.Add('       c.clave_prov,');
              D.IBQueryMicrosip.SQL.Add('       p.rfc_curp,');
              D.IBQueryMicrosip.SQL.Add('       l.permitir_sin_recepcion,');
              D.IBQueryMicrosip.SQL.Add('       l.pctje_rechazo,');
              D.IBQueryMicrosip.SQL.Add('       l.referencia');
              D.IBQueryMicrosip.SQL.Add('  FROM proveedores p');
              D.IBQueryMicrosip.SQL.Add('  JOIN claves_proveedores c ON (p.proveedor_id = c.proveedor_id)');
              D.IBQueryMicrosip.SQL.Add('  JOIN libres_proveedor l ON (p.proveedor_id = l.proveedor_id )');

              if (LAST <> '') then
                begin
                  D.IBQueryMicrosip.SQL.Add('WHERE p.FECHA_HORA_CREACION > ''' + LAST + '''');
                  D.IBQueryMicrosip.SQL.Add('   OR p.FECHA_HORA_ULT_MODIF > ''' + LAST + '''');
                  D.IBQueryMicrosip.SQL.Add('  AND c.rol_clave_prov_id = 49');
                end;

              D.IBQueryMicrosip.SQL.Add('ORDER BY p.proveedor_id');
              D.IBQueryMicrosip.Active := True;
              D.IBQueryMicrosip.First;
              while not D.IBQueryMicrosip.Eof do
                begin
                  Line := '"' + D.IBQueryMicrosip.FieldByName('PROVEEDOR_ID').AsString + '",';
                  Line := Line + '"' + StringReplace( StringReplace( D.IBQueryMicrosip.FieldByName('NOMBRE').AsString, '`', '''''', [rfReplaceAll] ), '''', '''''', [rfReplaceAll] ) + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('ESTATUS').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('CLAVE_PROV').AsString + '",';
                  Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryMicrosip.FieldByName('FECHA_HORA_ULT_MODIF').AsDateTime) + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('PCTJE_RECHAZO').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('REFERENCIA').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('RFC_CURP').AsString + '",';

                  if (D.IBQueryMicrosip.FieldByName('PERMITIR_SIN_RECEPCION').AsString = 'S') then
                    begin
                      Line := Line + '"SI",';
                    end
                  else
                    begin
                      Line := Line + '"NO",';
                    end;

                  Line := Line + '"' + Empresa_ID + '"';

                  List.Add(Line);
                  D.IBQueryMicrosip.Next;
                end;

              List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Proveedores');
              Total := Total + D.IBQueryMicrosip.RecordCount;

              {$ENDREGION}
            end;

          if actualiza_recepciones then
            begin
              {$REGION 'REVISA LA CANTIDAD DE RECEPCIONES A SUBIR O ACTUALIZAR'}

              List.Clear;

              if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Recepciones')) then
                begin
                  List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Recepciones');
                end
              else
                begin
                  List.Add('DOCTO_CM_ID,FOLIO,FECHA,CLAVE_PROV,PROVEEDOR_ID,MONEDA_ID,NOMBRE,CLAVE_FISCAL,IMPORTE_NETO,TOTAL_IMPUESTOS,TOTAL_RETENCIONES,FECHA_HORA_ULT_MODIF,ALMACEN_ID,ESTATUS,USO_CFDI,EMPRESA_ID,EMPRESA_NOMBRE');
                end;

              D.IBQueryMicrosip.Active := False;
              D.IBQueryMicrosip.SQL.Clear;

              D.IBQueryMicrosip.SQL.Add('SELECT' );
              // D.IBQueryMicrosip.SQL.Add('       dc.*, ' ); // SE QUITO POR QUE SE AGREGO CAMPO BOOLEANO EN MICROSIP 2024 GEN_X_CFDI
              D.IBQueryMicrosip.SQL.Add('       dc.DOCTO_CM_ID, ');
              D.IBQueryMicrosip.SQL.Add('       dc.FOLIO, ');
              D.IBQueryMicrosip.SQL.Add('       dc.FECHA, ');
              D.IBQueryMicrosip.SQL.Add('       dc.CLAVE_PROV, ');
              D.IBQueryMicrosip.SQL.Add('       dc.PROVEEDOR_ID, ');
              D.IBQueryMicrosip.SQL.Add('       dc.MONEDA_ID, ');
              D.IBQueryMicrosip.SQL.Add('       dc.IMPORTE_NETO, ');
              D.IBQueryMicrosip.SQL.Add('       dc.TOTAL_IMPUESTOS, ');
              D.IBQueryMicrosip.SQL.Add('       dc.TOTAL_RETENCIONES, ');
              D.IBQueryMicrosip.SQL.Add('       dc.FECHA_HORA_ULT_MODIF, ');
              D.IBQueryMicrosip.SQL.Add('       dc.ALMACEN_ID, ');
              D.IBQueryMicrosip.SQL.Add('       dc.ESTATUS, ');
              D.IBQueryMicrosip.SQL.Add('       m.NOMBRE, ');
              D.IBQueryMicrosip.SQL.Add('       m.CLAVE_FISCAL, ');
              D.IBQueryMicrosip.SQL.Add('       la.valor_desplegado USO_CFDI ');
              D.IBQueryMicrosip.SQL.Add('  FROM doctos_cm dc');
              D.IBQueryMicrosip.SQL.Add('  JOIN monedas m ON (dc.moneda_id = m.moneda_id)' );
              D.IBQueryMicrosip.SQL.Add('  LEFT JOIN libres_rec_cm lc ON(lc.docto_cm_id = dc.docto_cm_id)' );
              D.IBQueryMicrosip.SQL.Add('  LEFT JOIN listas_atributos la ON(lc.uso_cfdi = la.lista_atrib_id)' );
              D.IBQueryMicrosip.SQL.Add(' WHERE dc.tipo_docto = ''R''' );

              if (LAST <> '') then
                begin
                  D.IBQueryMicrosip.SQL.Add('AND (dc.FECHA_HORA_CREACION > ''' + LAST + ''' OR dc.FECHA_HORA_ULT_MODIF > ''' + LAST + ''')');
                end
              else
                begin
                  D.IBQueryMicrosip.SQL.Add('AND dc.fecha > ''01.03.2025''');
                end;

              D.IBQueryMicrosip.SQL.Add('ORDER BY dc.docto_cm_id');
              D.IBQueryMicrosip.Active := True;
              D.IBQueryMicrosip.First;
              while not D.IBQueryMicrosip.Eof do
                begin
                  Line := '"' + D.IBQueryMicrosip.FieldByName('DOCTO_CM_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('FOLIO').AsString + '",';
                  Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryMicrosip.FieldByName('FECHA').AsDateTime) + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('CLAVE_PROV').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('PROVEEDOR_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('MONEDA_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('NOMBRE').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('CLAVE_FISCAL').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('IMPORTE_NETO').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('TOTAL_IMPUESTOS').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('TOTAL_RETENCIONES').AsString + '",';
                  Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryMicrosip.FieldByName('FECHA_HORA_ULT_MODIF').AsDateTime) + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('ALMACEN_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('ESTATUS').AsString + '",';
                  Line := Line + '"' + Func.GET_USO_CLAVE(D.IBQueryMicrosip.FieldByName('USO_CFDI').AsString) + '",';
                  Line := Line + '"' + Empresa_ID + '",';
                  Line := Line + '"' + Empresa_N + '"';

                  List.Add(Line);
                  D.IBQueryMicrosip.Next;
                end;

              List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Recepciones');
              Total := Total + D.IBQueryMicrosip.RecordCount;

              {$ENDREGION}
            end;

          if actualiza_creditos then
            begin
              {$REGION 'REVISA LA CANTIDAD DE CREDITOS A SUBIR O ACTUALIZAR'}

              List.Clear;

              if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Creditos')) then
                begin
                  List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Creditos');
                end
              else
                begin
                  List.Add('DOCTO_CP_ID,CONCEPTO_CP_ID,CONCEPTO_CP,FOLIO,FECHA,CLAVE_PROV,PROVEEDOR_ID,DESCRIPCION,FECHA_HORA_ULT_MODIF,CANCELADO,APLICADO,TIENE_CFD,EMPRESA_ID,EMPRESA_NOMBRE');
                end;

              D.IBQueryMicrosip.Active := False;
              D.IBQueryMicrosip.SQL.Clear;
              D.IBQueryMicrosip.SQL.Add('SELECT');
              D.IBQueryMicrosip.SQL.Add('       dc.docto_cp_id,');
              D.IBQueryMicrosip.SQL.Add('       dc.concepto_cp_id,');
              D.IBQueryMicrosip.SQL.Add('       cc.nombre,');
              D.IBQueryMicrosip.SQL.Add('       dc.folio,');
              D.IBQueryMicrosip.SQL.Add('       dc.fecha,');
              D.IBQueryMicrosip.SQL.Add('       dc.clave_prov,');
              D.IBQueryMicrosip.SQL.Add('       dc.proveedor_id,');
              D.IBQueryMicrosip.SQL.Add('       dc.tipo_cambio,');
              D.IBQueryMicrosip.SQL.Add('       dc.cancelado,');
              D.IBQueryMicrosip.SQL.Add('       dc.aplicado,');
              D.IBQueryMicrosip.SQL.Add('       dc.descripcion,');
              D.IBQueryMicrosip.SQL.Add('       dc.tiene_cfd,');
              D.IBQueryMicrosip.SQL.Add('       dc.fecha_hora_ult_modif,'); // SE AGREGO LA ","
              D.IBQueryMicrosip.SQL.Add('       db.aplicado APLICADO_BA'); // SE AGREGO PARA VALIDAR PAGO LIBERADO
              D.IBQueryMicrosip.SQL.Add('  FROM doctos_cp dc');
              D.IBQueryMicrosip.SQL.Add('  JOIN conceptos_cp cc ON(dc.concepto_cp_id = cc.concepto_cp_id)');
              D.IBQueryMicrosip.SQL.Add('  LEFT JOIN doctos_entre_sis de ON(dc.docto_cp_id = de.docto_fte_id)'); // SE AGREGO PARA VALIDAR PAGO LIBERADO
              D.IBQueryMicrosip.SQL.Add('  LEFT JOIN doctos_ba db ON(de.docto_dest_id = db.docto_ba_id)'); // SE AGREGO PARA VALIDAR PAGO LIBERADO
              D.IBQueryMicrosip.SQL.Add(' WHERE dc.naturaleza_concepto = ''R''');
              D.IBQueryMicrosip.SQL.Add('   AND cc.tipo = ''P''');
              D.IBQueryMicrosip.SQL.Add('   AND (de.clave_sis_dest = ''BA'' OR de.clave_sis_dest IS NULL)'); // SE AGREGO PARA VALIDAR PAGO LIBERADO
              D.IBQueryMicrosip.SQL.Add('   AND (db.aplicado = ''S'' OR db.aplicado IS NULL)'); // SE AGREGO PARA VALIDAR PAGO LIBERADO

              { if (LAST <> '') then
                begin
                  D.IBQueryMicrosip.SQL.Add('AND (dc.FECHA_HORA_CREACION > ''' + LAST + ''' OR dc.FECHA_HORA_ULT_MODIF > ''' + LAST + ''')');
                end
              else
                begin
                  D.IBQueryMicrosip.SQL.Add('AND dc.fecha > ''01.03.2025''');
                end; }

              D.IBQueryMicrosip.SQL.Add('AND dc.fecha > ''01.01.2025''');

              D.IBQueryMicrosip.SQL.Add('ORDER BY dc.docto_cp_id');
              D.IBQueryMicrosip.Active := True;
              D.IBQueryMicrosip.First;
              while not D.IBQueryMicrosip.Eof do
                begin
                  // descripcion := D.IBQueryMicrosip.FieldByName('DESCRIPCION').AsString;
                  descripcion := StringReplace(D.IBQueryMicrosip.FieldByName('DESCRIPCION').AsString, '"', '', [rfReplaceAll]);

                  Line := '"' + D.IBQueryMicrosip.FieldByName('DOCTO_CP_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('CONCEPTO_CP_ID').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('NOMBRE').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('FOLIO').AsString + '",';
                  Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryMicrosip.FieldByName('FECHA').AsDateTime) + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('CLAVE_PROV').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('PROVEEDOR_ID').AsString + '",';
                  Line := Line + '"' + descripcion + '",';
                  // Line := Line + '"' + QuotedStr(descripcion) + '",';
                  Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.IBQueryMicrosip.FieldByName('FECHA_HORA_ULT_MODIF').AsDateTime) + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('CANCELADO').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('APLICADO').AsString + '",';
                  Line := Line + '"' + D.IBQueryMicrosip.FieldByName('TIENE_CFD').AsString + '",';
                  Line := Line + '"' + Empresa_ID + '",';
                  Line := Line + '"' + Empresa_N + '"';

                  List.Add(Line);
                  D.IBQueryMicrosip.Next;
                end;

              List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Creditos');
              Total := Total + D.IBQueryMicrosip.RecordCount;

              {$ENDREGION}
            end;

          if actualiza_particulares then
            begin
              {$REGION 'REVISA LA CANTIDAD DE PROVEEDORES EN EL PORTAL A REVISAR PARA ACTUALIZAR LOS DATOS PARTICULARES'}

              List.Clear;

              if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Libres')) then
                begin
                  List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Libres');
                end
              else
                begin
                  List.Add('PROVEEDOR_ID_MSP,CUENTA,SUCURSAL,CLABE,BANCO_NOMBRE,MAIL,REFERENCIA,NOMBRE,EMPRESA_ID,EMPRESA_NOMBRE');
                end;

              D.ADOQueryActual.Active := False;
              D.ADOQueryActual.SQL.Clear;
              D.ADOQueryActual.SQL.Add('SELECT');
              D.ADOQueryActual.SQL.Add('       PROVEEDOR_ID_MSP,');
              D.ADOQueryActual.SQL.Add('       NOMBRE,');
              D.ADOQueryActual.SQL.Add('       CUENTA,');
              D.ADOQueryActual.SQL.Add('       SUCURSAL,');
              D.ADOQueryActual.SQL.Add('       CLABE,');
              D.ADOQueryActual.SQL.Add('       BANCO_NOMBRE,');
              D.ADOQueryActual.SQL.Add('       MAIL,');
              D.ADOQueryActual.SQL.Add('       REFERENCIA,');
              D.ADOQueryActual.SQL.Add('       EMP_FK');
              D.ADOQueryActual.SQL.Add('  FROM PROVEEDORES_MSP ');
              D.ADOQueryActual.SQL.Add(' WHERE CUENTA IS NOT NULL');
              D.ADOQueryActual.SQL.Add('   AND EMP_FK = ' + Empresa_ID);
              D.ADOQueryActual.Active := True;
              D.ADOQueryActual.First;
              while not D.ADOQueryActual.Eof do
                begin
                  Line := '"' + D.ADOQueryActual.FieldByName('PROVEEDOR_ID_MSP').AsString + '",';
                  Line := Line + '"' + D.ADOQueryActual.FieldByName('CUENTA').AsString + '",';
                  Line := Line + '"' + D.ADOQueryActual.FieldByName('SUCURSAL').AsString + '",';
                  Line := Line + '"' + D.ADOQueryActual.FieldByName('CLABE').AsString + '",';
                  Line := Line + '"' + D.ADOQueryActual.FieldByName('BANCO_NOMBRE').AsString + '",';
                  Line := Line + '"' + D.ADOQueryActual.FieldByName('MAIL').AsString + '",';
                  Line := Line + '"' + D.ADOQueryActual.FieldByName('REFERENCIA').AsString + '",';
                  Line := Line + '"' + D.ADOQueryActual.FieldByName('NOMBRE').AsString + '",';
                  Line := Line + '"' + Empresa_ID + '",';
                  Line := Line + '"' + Empresa_N + '"';

                  List.Add( Line );
                  D.ADOQueryActual.Next;
                end;

              List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Libres');
              Total := Total + D.ADOQueryActual.RecordCount;

              {$ENDREGION}
            end;

          if actualiza_facturas then
            begin
              {$REGION 'REVISA FACTURAS A INSERTAR EN 3.3'}

              if (D.AplicaFacturas = True) then
                begin
                  List.Clear;

                  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Facturas_33')) then
                    begin
                      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Facturas_33');
                    end
                  else
                    begin
                      List.Add('DOCTO_CM_ID,FOLIO_COMPRA,IMPORTE_NETO,TOTAL_IMPUESTOS,TOTAL_RETENCIONES,DESCUENTO_GLOBAL,MONEDA_SIMBOLO,TIPO_CAMBIO,RECEPCION_ID,RECEP_ID,FOLIO_RECEPCION,FECHA_PAGO,FECHA_FACTURA,FECHA_RECEPCION,FECHA,PROVEEDOR_ID,RFC,NOMBRE,UUID,EMPRESA_ID,EMPRESA_NOMBRE');
                    end;

                  D.ADOQueryActual.Active := False;
                  D.ADOQueryActual.SQL.Clear;
                  D.ADOQueryActual.SQL.Add('SELECT');
                  D.ADOQueryActual.SQL.Add('       F.DOCTO_CM_ID,');
                  D.ADOQueryActual.SQL.Add('       F.FOLIO AS FOLIO_COMPRA,');
                  D.ADOQueryActual.SQL.Add('       F.IMPORTE_NETO,');
                  D.ADOQueryActual.SQL.Add('       F.TOTAL_IMPUESTOS,');
                  D.ADOQueryActual.SQL.Add('       F.TOTAL_RETENCIONES,');
                  D.ADOQueryActual.SQL.Add('       F.DESCUENTO_GLOBAL,');
                  D.ADOQueryActual.SQL.Add('       F.MONEDA_SIMBOLO,');
                  D.ADOQueryActual.SQL.Add('       F.TIPO_CAMBIO,');
                  D.ADOQueryActual.SQL.Add('       F.RECEPCION_ID,');
                  D.ADOQueryActual.SQL.Add('       F.RECEP_ID,');
                  D.ADOQueryActual.SQL.Add('       C.FOLIO AS FOLIO_RECEPCION,');
                  D.ADOQueryActual.SQL.Add('       F.FECHA_PAGO,');
                  D.ADOQueryActual.SQL.Add('       F.FECHA_FACTURA,');
                  D.ADOQueryActual.SQL.Add('       F.FECHA_RECEPCION,');
                  D.ADOQueryActual.SQL.Add('       C.FECHA,');
                  D.ADOQueryActual.SQL.Add('       F.PROVEEDOR_FK,');
                  D.ADOQueryActual.SQL.Add('       F.RFC,');
                  D.ADOQueryActual.SQL.Add('       P.NOMBRE,');
                  D.ADOQueryActual.SQL.Add('       F.UUID,');
                  D.ADOQueryActual.SQL.Add('       R.XML');
                  D.ADOQueryActual.SQL.Add('  FROM FACTURA_PROVEEDOR_33 F');
                  D.ADOQueryActual.SQL.Add(' INNER JOIN PROVEEDORES_MSP P ON((F.PROVEEDOR_FK = P.PROVEEDOR_ID_MSP) AND (F.EMP_FK = P.EMP_FK))');
                  D.ADOQueryActual.SQL.Add(' INNER JOIN ALMACENES_MSP A ON((A.ALMACEN_ID_MSP = F.ALMACEN_FK_MSP) AND (A.EMP_FK = F.EMP_FK))');
                  D.ADOQueryActual.SQL.Add(' INNER JOIN RECEPCIONES C ON((F.RECEP_ID = C.RECEP_ID) AND (F.EMP_FK = C.EMP_FK))');
                  D.ADOQueryActual.SQL.Add(' INNER JOIN ARCHIVOS_FACTURA_PROVEEDOR_33 R ON(F.UUID = R.UUID)');
                  D.ADOQueryActual.SQL.Add(' WHERE F.EMP_FK = ' + Empresa_ID);
                  D.ADOQueryActual.SQL.Add('   AND F.ESTATUS = ''S''');
                  D.ADOQueryActual.SQL.Add('   AND C.ESTATUS <> ''C''');
                  D.ADOQueryActual.SQL.Add('   AND F.RECEPCION_ID <> 0');
                  // D.ADOQueryActual.SQL.Add('   AND F.FECHA_RECEPCION > ''2025-07-01'' ');

                  // D.ADOQueryActual.SQL.Add('   AND F.FOLIO = ''A00032242'' ');

                  D.ADOQueryActual.SQL.Add(' ORDER BY P.NOMBRE, F.FECHA_PAGO ASC');

                  // D.ADOQueryActual.SQL.Add(' LIMIT 20');

                  D.ADOQueryActual.Active := True;
                  D.ADOQueryActual.First;
                  while not D.ADOQueryActual.Eof do
                    begin
                      if (D.ADOQueryActual.FieldByName('UUID').AsString <> '') then
                        begin
                          Line := '"' + D.ADOQueryActual.FieldByName('DOCTO_CM_ID').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('FOLIO_COMPRA').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('IMPORTE_NETO').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('TOTAL_IMPUESTOS').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('TOTAL_RETENCIONES').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('DESCUENTO_GLOBAL').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('MONEDA_SIMBOLO').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('TIPO_CAMBIO').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('RECEPCION_ID').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('RECEP_ID').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('FOLIO_RECEPCION').AsString + '",';
                          Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.ADOQueryActual.FieldByName('FECHA_PAGO').AsDateTime) + '",';
                          Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.ADOQueryActual.FieldByName('FECHA_FACTURA').AsDateTime) + '",';
                          Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.ADOQueryActual.FieldByName('FECHA_RECEPCION').AsDateTime) + '",';
                          Line := Line + '"' + FormatDateTime('DD/MM/YYYY HH:NN:SS', D.ADOQueryActual.FieldByName('FECHA').AsDateTime) + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('PROVEEDOR_FK').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('RFC').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('NOMBRE').AsString + '",';
                          Line := Line + '"' + D.ADOQueryActual.FieldByName('UUID').AsString + '",';
                          Line := Line + '"' + Empresa_ID + '",';
                          Line := Line + '"' + Empresa_N + '"';

                          // GUARDAMOS EL XML EN "Update/XML"
                          { XML := StringReplace(D.ADOQueryActual.FieldByName('XML').AsString, '<?xml version="1.0"?>', '<?xml version="1.0" encoding="UTF-8"?>', [rfReplaceAll]);
                          D.XML_FILE.Active := False;
                          D.XML_FILE.Options := [doNodeAutoCreate];
                          D.XML_FILE.LoadFromXML(XML);
                          D.XML_FILE.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/XML/' + D.ADOQueryActual.FieldByName('UUID').AsString + '.xml'); // }

                          List.Add(Line);
                        end
                      else
                        begin
                          EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Factura sin folio SAT ' + D.ADOQueryActual.FieldByName('FOLIO').AsString);
                        end;

                      D.ADOQueryActual.Next;
                    end;

                  List.SaveToFile(ExtractFilePath(ParamStr(0)) + '/Update/Facturas_33');
                  Total := Total + D.ADOQueryActual.RecordCount;
                end;

              {$ENDREGION}
            end;
        except
          on E : Exception do
            begin
              EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al revisar los registros de la empresa ' + D.ADOQueryEmpresas.FieldByName('EMP_NOMBRE').AsString);

              List.Destroy;
              D.XML_FILE.Active := False;
              Result := False;

              Exit;
            end;
        end;

        D.ADOQueryEmpresas.Next;
      end;
  except
    on E : Exception do
      begin
        EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al revisar los registros de las empresas autorizadas');

        List.Destroy;
        D.XML_FILE.Active := False;
        Result := False;

        Exit;
      end;
  end;

  List.Destroy;
  D.XML_FILE.Active := False;
  Result := True;
end;
{$ENDREGION}

Function CALCULA_REGISTROS(var Total :Integer):Boolean;
  var
    List :TStringList;
    R_Total :Integer;
begin
  Result := False; // INICIALIZA EN FALSO SI TERMINA EL PROCESO COMPLETO RETORNA VERDADERO

  List := TStringList.Create;

  Total := 0;

  {$REGION 'REVISAMOS SI NO QUEDO PENDIENTE NADA DE ACTUALIZAR DE ALGUNA PASADA ANTERIOR'}

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Empresas')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Empresas');
      Total := Total + (List.Count - 2);
    end;

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Almacenes')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Almacenes');
      Total := Total + (List.Count - 2);
    end;

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Monedas')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Monedas');
      Total := Total + (List.Count - 2);
    end;

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Proveedores')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Proveedores');
      Total := Total + (List.Count - 2);
    end;

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Recepciones')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Recepciones');
      Total := Total + (List.Count - 2);
    end;

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Creditos')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Creditos');
      Total := Total + (List.Count - 2);
    end;

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Libres')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Libres');
      Total := Total + (List.Count - 2);
    end;

  List.Clear;
  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Facturas_33')) then
    begin
      List.LoadFromFile(ExtractFilePath(ParamStr(0)) + '/Update/Facturas_33');
      Total := Total + (List.Count - 2);
    end;

  {$ENDREGION}

  if (CALCULA_EMPRESAS(R_Total) = False) then
    begin
      D.Conexion_MySQL.Connected := False;
      D.Conexion_Config.Connected := False;
      D.Conexion_Microsip.Connected := False;

      Exit;
    end;
  Total := Total + R_Total;

  if (CALCULA_REGISTROS_EMPRESAS(R_Total) = False) then
    begin
      D.Conexion_MySQL.Connected := False;
      D.Conexion_Config.Connected := False;
      D.Conexion_Microsip.Connected := False;

      Exit;
    end;
  Total := Total + R_Total;

  D.Conexion_MySQL.Connected := False;
  D.Conexion_Config.Connected := False;
  D.Conexion_Microsip.Connected := False;

  Result := True;
end;

end.
