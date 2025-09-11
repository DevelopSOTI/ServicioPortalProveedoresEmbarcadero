unit Func_Facturas_3_2;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf;

  // INSERCIÓN DE FACTURAS 3.2 DESCONTINUADO YA NO SE USARA
  Function SELECT_FACTURAS_APLICAR():Boolean;

implementation

uses
  Data, Func;

{$REGION 'ACTUALIZAR_FACTURA_PORTAL - FUNCIÓN QUE ACTUALIZA LAS RECEPCIONES Y FACTURAS EN MYSQL'}
Function ACTUALIZAR_FACTURA_PORTAL( FOLIO_COMPRA, FOLIO_RECEPCION, DOCTO_CM_ID, RECEP_ID :string ):Boolean;
  var
    CadenaSQL :string;
begin
  // CAMBIO LOS ESTATUS EN LAS FACTURAS EN MYSQL
  try
    CadenaSQL := CadenaSQL + 'UPDATE FACTURA_PROVEEDOR SET';
    CadenaSQL := CadenaSQL + '       FOLIO_MSP = ' + '''' + FOLIO_COMPRA + ''',';
    CadenaSQL := CadenaSQL + '       COMPRA_ID = ' + DOCTO_CM_ID + ',';
    CadenaSQL := CadenaSQL + '       ESTATUS = ''R'',';
    CadenaSQL := CadenaSQL + '       USUARIO_CONV_COMPRA = ' + '''' + 'SYSDBA' + ''',';
    CadenaSQL := CadenaSQL + '       FECHA_CONV_COMPRA = ' + '''' + FormatDateTime( 'YYYY-MM-DD hh:nn:ss', Now ) + '''';
    CadenaSQL := CadenaSQL + ' WHERE RECEP_ID = ' + RECEP_ID;

    D.MySQL_Command.CommandText := CadenaSQL;
    D.MySQL_Command.Execute;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo actualizar el estatus de la factura ' + FOLIO_COMPRA + ' en el portal' );
        D.Transaction_Microsip.RollbackRetaining;
        Result := False;
        Exit;
      end;
  end;

  // ACTUALIZA RECEPCION EN EL PORTAL A RECIBIDA
  try
    D.MySQL_Command.CommandText := 'UPDATE RECEPCIONES SET ESTATUS = ''R'' WHERE RECEP_ID = ' + RECEP_ID;
    D.MySQL_Command.Execute;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo actualizar el estatus de la recepcion ' + FOLIO_RECEPCION + ' en el portal web' );
        D.Transaction_Microsip.RollbackRetaining;
        Result := False;
        Exit;
      end;
  end;

  Result := True;
end;
{$ENDREGION}

{$REGION 'APLICAR_MICROSIP'}
procedure APLICAR_MICROSIP( DOCTO_CM_ID_MYSQL, RECEP_ID, RECEPCION_ID, EMPRESA_ID, FOLIO_RECEPCION, FOLIO_COMPRA, UUID, RFC, NOMBRE, MONEDA_SIMBOLO :String; PROVEEDOR_ID :Integer; FECHA_PAGO, FECHA_FACTURA, FECHA_RECEPCION, FECHA :TDateTime; IMPORTE_NETO, TOTAL_IMPUESTOS, TOTAL_RETENCIONES, DESCUENTO_GLOBAL, TIPO_CAMBIO :Double; var LINE :string );
  var
    CadenaSQL, RzEditFolioFac, RzEditFecha_fac, RzEditPecha_prov, FOLIO_FINAL, FOLIO_XML, XML :String;
    DOCTO_CM_ID, DOCTO_CM_DET_ID, DOCTO_CM_LIGA_ID, CFDI_ID, DOCTO_CP_ID, IMPTE_DOCTO_CP_ID :Integer;
begin
  {$REGION 'BUSCO LOS DATOS DEL ENCABEZADO DE LA RECEPCIÓN EN MICROSIP POR FOLIO Y PROVEEDOR ( DOCTO_CM_ID, PROVEEDOR_ID )'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    D.IBQueryMicrosip.SQL.Add( 'SELECT * FROM DOCTOS_CM' );
    D.IBQueryMicrosip.SQL.Add( ' WHERE FOLIO = ''' + FOLIO_RECEPCION + '''' );
    D.IBQueryMicrosip.SQL.Add( '   AND PROVEEDOR_ID = ' + IntToStr( PROVEEDOR_ID ) );
    D.IBQueryMicrosip.SQL.Add( '   AND TIPO_DOCTO = ''R''' );
    D.IBQueryMicrosip.Active := True;
    D.IBQueryMicrosip.Last;

    {$REGION 'SI NO HUBO RENGLONES ENTONCES NO EXISTE LA RECEPCIÓN Y SE SALE DEL PROCESO'}
    if ( D.IBQueryMicrosip.RecordCount = 0 ) then
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', 'No se puede insertar la compra porque no exise la recepción ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        D.IBQueryMicrosip.Active := False;
        Exit;
      end;
    {$ENDREGION}

    DOCTO_CM_ID := D.IBQueryMicrosip.FieldByName('DOCTO_CM_ID').AsInteger;

    {$REGION 'REVISA EL ESTATUS DE LA RECEPCIÓN, QUE NO ESTE FACTURADA'}
    if ( D.IBQueryMicrosip.FieldByName( 'ESTATUS' ).AsString = 'F' ) then
      begin
        try
          D.SELECT.Active := False;
          D.SELECT.SQL.Clear;
          D.SELECT.SQL.Add( 'SELECT * FROM doctos_cm_ligas d' );
          D.SELECT.SQL.Add( '  JOIN doctos_cm dc ON ( d.docto_cm_dest_id = dc.docto_cm_id )' );
          D.SELECT.SQL.Add( ' WHERE d.docto_cm_fte_id = ' + IntToStr( DOCTO_CM_ID ) );
          D.SELECT.Active := True;

          DOCTO_CM_ID := D.SELECT.FieldByName('DOCTO_CM_DEST_ID').AsInteger;
          FOLIO_COMPRA := D.SELECT.FieldByName('FOLIO').AsString;

          // EN ESTA PARTE USA FOLIO_COMPRA EN VEZ DE FOLIO_FINAL PORQUE ES UN FOLIO QUE YA ESTA REGISTRADO
          ACTUALIZAR_FACTURA_PORTAL( FOLIO_COMPRA, FOLIO_RECEPCION, IntToStr( DOCTO_CM_ID ), RECEP_ID );

          EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', 'No se puede insertar la compra porque la recepción ' + FOLIO_RECEPCION + ' ya fue facturada' );
          D.Transaction_Microsip.RollbackRetaining;
          D.SELECT.Active := False;
          Exit;
        except
          on E : Exception do
            begin
              EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al intentar cargar el identificador de la compra de la recepción ' + FOLIO_RECEPCION );
              D.Transaction_Microsip.RollbackRetaining;
              D.IBQueryMicrosip.Active := False;
              Exit;
            end;
        end;
      end;
    {$ENDREGION}
  except
    on E:Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al cargar los datos de Microsip (Recepción)' );
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'BUSCA QUE NO ESTE REGISTRADA LA FACTURA CON ESTE PROVEEDOR, SI VIENE EN 0s QUIERE DECIR QUE EL XML NO TENIA FOLIO Y NO ENTRA'}
  if ( FOLIO_COMPRA <> '000000000' ) then
    begin
      try
        D.SELECT.Active := False;
        D.SELECT.SQL.Clear;
        D.SELECT.SQL.Add('SELECT * FROM DOCTOS_CM' );
        D.SELECT.SQL.Add(' WHERE FOLIO_PROV = ''' + FOLIO_COMPRA + '''');
        D.SELECT.SQL.Add('   AND PROVEEDOR_ID = ' + IntToStr( PROVEEDOR_ID ));
        D.SELECT.SQL.Add('   AND TIPO_DOCTO = ''C''');
        D.SELECT.Active := True;
        D.SELECT.Last;

        if ( D.SELECT.RecordCount <> 0 ) then
          begin
            EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', 'No se puede insertar porque el folio de la factura ' + FOLIO_COMPRA + ' ya se encuentra registrado' );
            D.Transaction_Microsip.RollbackRetaining;
            D.IBQueryMicrosip.Active := False;
            D.SELECT.Active := False;
            Exit;
          end;
      except
        on E:Exception do
          begin
            EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al revisar si la factura ' + FOLIO_COMPRA + ' ya se encontraba registrada' );
            D.Transaction_Microsip.RollbackRetaining;
            D.IBQueryMicrosip.Active := False;
            Exit;
          end;
      end;
    end;
  {$ENDREGION}

  // SI PASO TODAS LAS VALIDACIONES, ENTONCES PROCEDE A INSERTAR CON EL FOLIO NUEVO
  FOLIO_FINAL := SIGUIENTE_FOLIO( 'WEB' );

  if ( FOLIO_FINAL = 'WEB000000' ) then
    begin
      EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', 'La serie WEB no esta registrada en la empresa ' + NOMBRE );
      D.Transaction_Microsip.RollbackRetaining;
      D.IBQueryMicrosip.Active := False;
      D.SELECT.Active := False;
      Exit;
    end;

  {$REGION 'SI VIENE EN 0s QUIERE DECIR QUE EL XML NO TENIA FOLIO Y NO ENTRA'}
  if ( FOLIO_COMPRA <> '000000000' ) then
    begin
      FOLIO_XML := FOLIO_COMPRA;
    end
  else
    begin
      FOLIO_COMPRA := FOLIO_FINAL;
      FOLIO_XML := '';
    end;
  {$ENDREGION}

  RzEditFolioFac := FOLIO_COMPRA;
  RzEditFecha_fac := FormatDateTime( 'dd/mm/yyyy', FECHA_FACTURA );
  RzEditPecha_prov := FormatDateTime( 'dd/mm/yyyy', FECHA_RECEPCION );

  {$REGION 'COLOCA EL ENCABEZADO DE LA COMPRA ( DOCTOS_CM )'}
  try
    D.GEN_DOCTO_ID.Prepare;
    D.GEN_DOCTO_ID.ExecProc;
    DOCTO_CM_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

    D.DOCTOS_CM.TableName := 'DOCTOS_CM';
    D.DOCTOS_CM.Active := True;
    D.DOCTOS_CM.Insert;

    D.DOCTOS_CM.FieldByName('DOCTO_CM_ID').AsInteger := DOCTO_CM_ID;
    D.DOCTOS_CM.FieldByName('TIPO_DOCTO').AsString := 'C';
    D.DOCTOS_CM.FieldByName('SUBTIPO_DOCTO').AsString := D.IBQueryMicrosip.FieldByName('SUBTIPO_DOCTO').AsString;
    D.DOCTOS_CM.FieldByName('FOLIO').AsString := FOLIO_FINAL;
    D.DOCTOS_CM.FieldByName('FECHA').AsDateTime := FECHA;
    D.DOCTOS_CM.FieldByName('CLAVE_PROV').AsString := D.IBQueryMicrosip.FieldByName('CLAVE_PROV').AsString;
    D.DOCTOS_CM.FieldByName('PROVEEDOR_ID').AsInteger := D.IBQueryMicrosip.FieldByName('PROVEEDOR_ID').AsInteger;
    D.DOCTOS_CM.FieldByName('FOLIO_PROV').AsString := FOLIO_COMPRA;
    D.DOCTOS_CM.FieldByName('FACTURA_DEV').AsString := '';
    D.DOCTOS_CM.FieldByName('ALMACEN_ID').AsInteger := D.IBQueryMicrosip.FieldByName('ALMACEN_ID').AsInteger;
    D.DOCTOS_CM.FieldByName('MONEDA_ID').AsInteger := D.IBQueryMicrosip.FieldByName('MONEDA_ID').AsInteger;
    D.DOCTOS_CM.FieldByName('TIPO_CAMBIO').AsFloat := D.IBQueryMicrosip.FieldByName('TIPO_CAMBIO').AsFloat;
    D.DOCTOS_CM.FieldByName('TIPO_DSCTO').AsString := D.IBQueryMicrosip.FieldByName('TIPO_DSCTO').AsString;
    D.DOCTOS_CM.FieldByName('DSCTO_PCTJE').AsFloat := D.IBQueryMicrosip.FieldByName('DSCTO_PCTJE').AsFloat;
    D.DOCTOS_CM.FieldByName('DSCTO_IMPORTE').AsFloat := D.IBQueryMicrosip.FieldByName('DSCTO_IMPORTE').AsFloat;
    D.DOCTOS_CM.FieldByName('ESTATUS').AsString := 'N';
    D.DOCTOS_CM.FieldByName('APLICADO').AsString := 'S';
    D.DOCTOS_CM.FieldByName('DESCRIPCION').AsString := D.IBQueryMicrosip.FieldByName('DESCRIPCION').AsString;
    D.DOCTOS_CM.FieldByName('IMPORTE_NETO').AsFloat := D.IBQueryMicrosip.FieldByName('IMPORTE_NETO').AsFloat;
    D.DOCTOS_CM.FieldByName('FLETES').AsFloat := D.IBQueryMicrosip.FieldByName('FLETES').AsFloat;
    D.DOCTOS_CM.FieldByName('OTROS_CARGOS').AsFloat := D.IBQueryMicrosip.FieldByName('OTROS_CARGOS').AsFloat;
    D.DOCTOS_CM.FieldByName('TOTAL_IMPUESTOS').AsFloat := D.IBQueryMicrosip.FieldByName('TOTAL_IMPUESTOS').AsFloat;
    D.DOCTOS_CM.FieldByName('TOTAL_RETENCIONES').AsFloat := D.IBQueryMicrosip.FieldByName('TOTAL_RETENCIONES').AsFloat;
    D.DOCTOS_CM.FieldByName('GASTOS_ADUANALES').AsFloat := D.IBQueryMicrosip.FieldByName('GASTOS_ADUANALES').AsFloat;
    D.DOCTOS_CM.FieldByName('OTROS_GASTOS').AsFloat := D.IBQueryMicrosip.FieldByName('OTROS_GASTOS').AsFloat;
    D.DOCTOS_CM.FieldByName('FORMA_EMITIDA').AsString := 'N';
    D.DOCTOS_CM.FieldByName('CONTABILIZADO').AsString := 'N';
    D.DOCTOS_CM.FieldByName('ACREDITAR_CXP').AsString := 'N';
    D.DOCTOS_CM.FieldByName('SISTEMA_ORIGEN').AsString := 'CM';
    D.DOCTOS_CM.FieldByName('COND_PAGO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('COND_PAGO_ID').AsInteger;
    D.DOCTOS_CM.FieldByName('PCTJE_DSCTO_PPAG').AsFloat := 0;
    D.DOCTOS_CM.FieldByName('CARGAR_SUN').AsString := D.IBQueryMicrosip.FieldByName('CARGAR_SUN').AsString;
    D.DOCTOS_CM.FieldByName('ENVIADO').AsString := 'N';
    D.DOCTOS_CM.FieldByName('TIENE_CFD').AsString := 'N'; // D.DOCTOS_CM.FieldByName('TIENE_CFD').AsString := 'S';
    D.DOCTOS_CM.FieldByName('USUARIO_CREADOR').AsString := 'SISTEMAWEB';
    D.DOCTOS_CM.FieldByName('USUARIO_AUT_CREACION').AsString := 'SYSDBA';
    D.DOCTOS_CM.FieldByName('USUARIO_ULT_MODIF').AsString := 'SISTEMAWEB';
    D.DOCTOS_CM.FieldByName('USUARIO_AUT_MODIF').AsString := 'SISTEMAWEB';

    if ( D.IBQueryMicrosip.FieldByName('CONSIG_CM_ID').AsInteger <> 0 ) then
      begin
        D.DOCTOS_CM.FieldByName('CONSIG_CM_ID').AsInteger := D.IBQueryMicrosip.FieldByName('CONSIG_CM_ID').AsInteger;
      end;

    if ( D.IBQueryMicrosip.FieldByName('PEDIMENTO_ID').AsInteger <> 0 ) then
      begin
        D.DOCTOS_CM.FieldByName('PEDIMENTO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('PEDIMENTO_ID').AsInteger;
      end;

    if ( D.IBQueryMicrosip.FieldByName('VIA_EMBARQUE_ID').AsInteger <> 0 ) then
      begin
        D.DOCTOS_CM.FieldByName('VIA_EMBARQUE_ID').AsInteger := D.IBQueryMicrosip.FieldByName('VIA_EMBARQUE_ID').AsInteger;
      end;

    if ( D.IBQueryMicrosip.FieldByName('IMPUESTO_SUSTITUIDO_ID').AsInteger <> 0 ) then
      begin
        D.DOCTOS_CM.FieldByName('IMPUESTO_SUSTITUIDO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('IMPUESTO_SUSTITUIDO_ID').AsInteger;
      end;

    if ( D.IBQueryMicrosip.FieldByName('IMPUESTO_SUSTITUTO_ID').AsInteger <> 0 ) then
      begin
        D.DOCTOS_CM.FieldByName('IMPUESTO_SUSTITUTO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('IMPUESTO_SUSTITUTO_ID').AsInteger;
      end;

    D.DOCTOS_CM.Post;
    D.DOCTOS_CM.Active := False;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el encabezado del documento de compra ' + FOLIO_COMPRA );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'COLOCA EL DOCTOS_CM_LIGAS ( DOCTOS_CM_LIGAS ) ( ENTRE RECEPCIÓN Y LA COMPRA CREADA )'}
  try
    D.GEN_DOCTO_ID.Prepare;
    D.GEN_DOCTO_ID.ExecProc;
    DOCTO_CM_LIGA_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

    D.DOCTOS_CM_LIGAS.TableName := 'DOCTOS_CM_LIGAS';
    D.DOCTOS_CM_LIGAS.Active := True;
    D.DOCTOS_CM_LIGAS.Insert;

    D.DOCTOS_CM_LIGAS.FieldByName('DOCTO_CM_LIGA_ID').AsInteger := DOCTO_CM_LIGA_ID;
    D.DOCTOS_CM_LIGAS.FieldByName('DOCTO_CM_FTE_ID').AsInteger := D.IBQueryMicrosip.FieldByName('DOCTO_CM_ID').AsInteger;
    D.DOCTOS_CM_LIGAS.FieldByName('DOCTO_CM_DEST_ID').AsInteger := DOCTO_CM_ID;

    D.DOCTOS_CM_LIGAS.Post;
    D.DOCTOS_CM_LIGAS.Active := False;
  except
    on E:Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el encabezado de la liga de la rececpción ' + FOLIO_RECEPCION + ' a la compra ' + FOLIO_COMPRA );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'COLOCA EL DETALLE DE LA COMPRA ( DOCTOS_CM_DET y DOCTOS_CM_LIGAS_DET )'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;

    D.IBQueryMicrosip.SQL.Add('SELECT * FROM DOCTOS_CM_DET D' );
    D.IBQueryMicrosip.SQL.Add('  JOIN DOCTOS_CM E ON ( D.DOCTO_CM_ID = E.DOCTO_CM_ID )');
    D.IBQueryMicrosip.SQL.Add(' WHERE E.FOLIO = ''' + FOLIO_RECEPCION + ''' AND E.TIPO_DOCTO = ''R''');
    D.IBQueryMicrosip.SQL.Add(' ORDER BY D.POSICION');

    D.IBQueryMicrosip.Active := True;
    D.IBQueryMicrosip.First;
    while not D.IBQueryMicrosip.Eof do
      begin
        D.GEN_DOCTO_ID.Prepare;
        D.GEN_DOCTO_ID.ExecProc;
        DOCTO_CM_DET_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

        {$REGION 'COLOCA EL DETALLE DE LA COMPRA'}
        try
          D.DOCTOS_CM_DET.TableName := 'DOCTOS_CM_DET';
          D.DOCTOS_CM_DET.Active := True;
          D.DOCTOS_CM_DET.Insert;

          D.DOCTOS_CM_DET.FieldByName('DOCTO_CM_DET_ID').AsInteger := DOCTO_CM_DET_ID;
          D.DOCTOS_CM_DET.FieldByName('DOCTO_CM_ID').AsInteger := DOCTO_CM_ID;
          D.DOCTOS_CM_DET.FieldByName('CLAVE_ARTICULO').AsString := D.IBQueryMicrosip.FieldByName('CLAVE_ARTICULO').AsString;
          D.DOCTOS_CM_DET.FieldByName('ARTICULO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('ARTICULO_ID').AsInteger;
          D.DOCTOS_CM_DET.FieldByName('UNIDADES').AsFloat := D.IBQueryMicrosip.FieldByName('UNIDADES').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('UNIDADES_REC_DEV').AsFloat := D.IBQueryMicrosip.FieldByName('UNIDADES_REC_DEV').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('UNIDADES_A_REC').AsFloat := D.IBQueryMicrosip.FieldByName('UNIDADES_A_REC').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('UMED').AsString := D.IBQueryMicrosip.FieldByName('UMED').AsString;
          D.DOCTOS_CM_DET.FieldByName('CONTENIDO_UMED').AsFloat := D.IBQueryMicrosip.FieldByName('CONTENIDO_UMED').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('PRECIO_UNITARIO').AsFloat := D.IBQueryMicrosip.FieldByName('PRECIO_UNITARIO').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('PCTJE_DSCTO').AsFloat := D.IBQueryMicrosip.FieldByName('PCTJE_DSCTO').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('PCTJE_DSCTO_PRO').AsFloat := D.IBQueryMicrosip.FieldByName('PCTJE_DSCTO_PRO').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('PCTJE_DSCTO_VOL').AsFloat := D.IBQueryMicrosip.FieldByName('PCTJE_DSCTO_VOL').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('PCTJE_DSCTO_PROMO').AsFloat := D.IBQueryMicrosip.FieldByName('PCTJE_DSCTO_PROMO').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('DSCTO_ART').AsFloat := D.IBQueryMicrosip.FieldByName('DSCTO_ART').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('DSCTO_EXTRA').AsFloat := D.IBQueryMicrosip.FieldByName('DSCTO_EXTRA').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('PRECIO_TOTAL_NETO').AsFloat := D.IBQueryMicrosip.FieldByName('PRECIO_TOTAL_NETO').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('PCTJE_ARANCEL').AsFloat := D.IBQueryMicrosip.FieldByName('PCTJE_ARANCEL').AsFloat;
          D.DOCTOS_CM_DET.FieldByName('NOTAS').AsString := D.IBQueryMicrosip.FieldByName('NOTAS').AsString;
          D.DOCTOS_CM_DET.FieldByName('POSICION').AsInteger := D.IBQueryMicrosip.FieldByName('POSICION').AsInteger;

          D.DOCTOS_CM_DET.Post;
          D.DOCTOS_CM_DET.Active := False;
        except
          on E : Exception do
            begin
              EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el detalle del documento de compra ' + FOLIO_COMPRA );
              D.Transaction_Microsip.RollbackRetaining;
              Exit;
            end;
        end;
        {$ENDREGION}

        {$REGION 'COLOCA EL DOCTOS_CM_LIGAS_DETALLE'}
        try
          D.DOCTOS_CM_LIGAS_DET.TableName := 'DOCTOS_CM_LIGAS_DET';
          D.DOCTOS_CM_LIGAS_DET.Active := True;
          D.DOCTOS_CM_LIGAS_DET.Insert;

          D.DOCTOS_CM_LIGAS_DET.FieldByName('DOCTO_CM_LIGA_ID').AsInteger := DOCTO_CM_LIGA_ID;
          D.DOCTOS_CM_LIGAS_DET.FieldByName('DOCTO_CM_DET_FTE_ID').AsInteger := D.IBQueryMicrosip.FieldByName('DOCTO_CM_DET_ID').AsInteger;
          D.DOCTOS_CM_LIGAS_DET.FieldByName('DOCTO_CM_DET_DEST_ID').AsInteger := DOCTO_CM_DET_ID;

          D.DOCTOS_CM_LIGAS_DET.Post;
          D.DOCTOS_CM_LIGAS_DET.Active := False;
        except
          on E : Exception do
            begin
              EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el detalle de la liga de la recepción ' + FOLIO_RECEPCION + ' con la compra ' + FOLIO_COMPRA );
              D.Transaction_Microsip.RollbackRetaining;
              Exit;
            end;
        end;
        {$ENDREGION}

        D.IBQueryMicrosip.Next;
      end;
  except
    on E:Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo cargar el detalle de la rececpción ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'COLOCA LOS IMPUESTOS DE LA COMPRA ( IMPUESTOS_DOCTOS_CM )'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    D.IBQueryMicrosip.SQL.Add( 'SELECT * FROM IMPUESTOS_DOCTOS_CM D' );
    D.IBQueryMicrosip.SQL.Add( '  JOIN DOCTOS_CM E ON( D.DOCTO_CM_ID = E.DOCTO_CM_ID )' );
    D.IBQueryMicrosip.SQL.Add( ' WHERE E.FOLIO = ''' + FOLIO_RECEPCION + ''' AND E.TIPO_DOCTO = ''R''' );
    D.IBQueryMicrosip.Active := True;
    D.IBQueryMicrosip.First;
    while not D.IBQueryMicrosip.Eof do
      begin
        {$REGION 'COLOCA EL IMPUESTOS_DOCTOS_CM'}
        try
          D.IMPUESTOS_DOCTOS_CM.TableName := 'IMPUESTOS_DOCTOS_CM';
          D.IMPUESTOS_DOCTOS_CM.Active := True;
          D.IMPUESTOS_DOCTOS_CM.Insert;

          D.IMPUESTOS_DOCTOS_CM.FieldByName('DOCTO_CM_ID').AsInteger := DOCTO_CM_ID;
          D.IMPUESTOS_DOCTOS_CM.FieldByName('IMPUESTO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('IMPUESTO_ID').AsInteger;
          D.IMPUESTOS_DOCTOS_CM.FieldByName('COMPRA_NETA').AsFloat := D.IBQueryMicrosip.FieldByName('COMPRA_NETA').AsFloat;
          D.IMPUESTOS_DOCTOS_CM.FieldByName('OTROS_IMPUESTOS').AsFloat := D.IBQueryMicrosip.FieldByName('OTROS_IMPUESTOS').AsFloat;
          D.IMPUESTOS_DOCTOS_CM.FieldByName('PCTJE_IMPUESTO').AsFloat := D.IBQueryMicrosip.FieldByName('PCTJE_IMPUESTO').AsFloat;
          D.IMPUESTOS_DOCTOS_CM.FieldByName('IMPORTE_IMPUESTO').AsFloat := D.IBQueryMicrosip.FieldByName('IMPORTE_IMPUESTO').AsFloat;
          D.IMPUESTOS_DOCTOS_CM.FieldByName('UNIDADES_IMPUESTO').AsFloat := D.IBQueryMicrosip.FieldByName('UNIDADES_IMPUESTO').AsFloat;
          D.IMPUESTOS_DOCTOS_CM.FieldByName('IMPORTE_UNITARIO_IMPUESTO').AsFloat := D.IBQueryMicrosip.FieldByName('IMPORTE_UNITARIO_IMPUESTO').AsFloat;

          D.IMPUESTOS_DOCTOS_CM.Post;
          D.IMPUESTOS_DOCTOS_CM.Active := False;
        except
          on E : Exception do
            begin
              EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el detalle de impuestos del documento de compra ' + FOLIO_COMPRA );
              D.Transaction_Microsip.RollbackRetaining;
              Exit;
            end;
        end;
        {$ENDREGION}

        D.IBQueryMicrosip.Next;
      end;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo cargar el detalle de impuestos de la recepción ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  LINE := '"' + UUID + '",';
  LINE := LINE + '"' + FOLIO_XML + '",';
  LINE := LINE + '"' + FormatDateTime( 'dd/mm/yyyy', FECHA_FACTURA ) + '",';
  LINE := LINE + '"' + RFC + '",';
  LINE := LINE + '"' + NOMBRE + '",';
  LINE := LINE + '"' + FloatToStr( IMPORTE_NETO + TOTAL_IMPUESTOS - TOTAL_RETENCIONES - DESCUENTO_GLOBAL ) + '",';
  LINE := LINE + '"' + MONEDA_SIMBOLO + '",';
  LINE := LINE + '"' + FloatToStr( TIPO_CAMBIO ) + '",';
  LINE := LINE + '"' + RFC + '_' + FOLIO_COMPRA + '.xml",';
  LINE := LINE + '"' + FOLIO_COMPRA + '",';
  LINE := LINE + '"' + IntToStr( DOCTO_CM_ID ) + '"';



  {$REGION '******************* INSERCION DE XML, COMENTADO *******************'}
  // COLOCA/BUSCA EL REPOSITORIO_CFDI - CARGA LOS ARCHIVOS QUE ESTAN EN EL PORTAL DE LA FACTURA EN PROCESO
  { try
    D.XML_FILE.Active := False;
    D.XML_FILE.Options := [doNodeAutoIndent];
    D.XML_FILE.LoadFromFile( ExtractFilePath( ParamStr( 0 ) ) + '/Update/XML/' + UUID + '.xml' );

    XML := xmlDoc.FormatXMLData( D.XML_FILE.XML.Text );
    XML := StringReplace( XML, '<?xml version="1.0"?>', '<?xml version="1.0" encoding="UTF-8"?>', [rfReplaceAll] );
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudieron cargar los archivos del portal de la compra ' + FOLIO_COMPRA );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end; }

  // BUSCA SI ESTA EL REPOSITORIO_CFDI
  { try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    D.IBQueryMicrosip.SQL.Add( 'SELECT * FROM REPOSITORIO_CFDI WHERE UUID = ''' + UUID + '''');
    D.IBQueryMicrosip.Active := True;
    D.IBQueryMicrosip.Last;

    // SI NO ESTA EL REPOSITORIO HAY QUE CREARLO
    if ( D.IBQueryMicrosip.RecordCount = 0 ) then
      begin
        try
          D.GEN_DOCTO_ID.Prepare;
          D.GEN_DOCTO_ID.ExecProc;
          CFDI_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

          D.REPOSITORIO_CFDI.Active := True;
          D.REPOSITORIO_CFDI.Insert;

          D.REPOSITORIO_CFDI.FieldByName('CFDI_ID').AsInteger := CFDI_ID;
          D.REPOSITORIO_CFDI.FieldByName('MODALIDAD_FACTURACION').AsString := 'CFDI';
          D.REPOSITORIO_CFDI.FieldByName('VERSION').AsString := '3.2';
          D.REPOSITORIO_CFDI.FieldByName('UUID').AsString := UUID;
          D.REPOSITORIO_CFDI.FieldByName('NATURALEZA').AsString := 'R';
          D.REPOSITORIO_CFDI.FieldByName('TIPO_COMPROBANTE').AsString := 'I';
          D.REPOSITORIO_CFDI.FieldByName('TIPO_DOCTO_MSP').AsString := 'Compra';
          D.REPOSITORIO_CFDI.FieldByName('FOLIO').AsString := FOLIO_XML;
          D.REPOSITORIO_CFDI.FieldByName('FECHA').AsDateTime := FECHA_FACTURA;
          D.REPOSITORIO_CFDI.FieldByName('RFC').AsString := RFC;
          D.REPOSITORIO_CFDI.FieldByName('NOMBRE').AsString := NOMBRE;
          D.REPOSITORIO_CFDI.FieldByName('IMPORTE').AsFloat := IMPORTE_NETO + TOTAL_IMPUESTOS - TOTAL_RETENCIONES - DESCUENTO_GLOBAL;
          D.REPOSITORIO_CFDI.FieldByName('MONEDA').AsString := MONEDA_SIMBOLO;
          D.REPOSITORIO_CFDI.FieldByName('TIPO_CAMBIO').AsFloat := TIPO_CAMBIO;
          D.REPOSITORIO_CFDI.FieldByName('ES_PARCIALIDAD').AsString := 'N';
          D.REPOSITORIO_CFDI.FieldByName('NOM_ARCH').AsString := RFC + '_' + FOLIO_COMPRA + '.xml';
          D.REPOSITORIO_CFDI.FieldByName('XML').AsString := XML;
          D.REPOSITORIO_CFDI.FieldByName('REFER_GRUPO').AsString := FOLIO_COMPRA;
          D.REPOSITORIO_CFDI.FieldByName('SELLO_VALIDADO').AsString := 'M';
          D.REPOSITORIO_CFDI.FieldByName('USUARIO_CREADOR').AsString := 'SISTEMAWEB';
          D.REPOSITORIO_CFDI.FieldByName('FECHA_HORA_CREACION').AsDateTime := Now;

          D.REPOSITORIO_CFDI.Post;
        except
          on E : Exception do
            begin
              EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el repositorio del CFDI ' + FOLIO_RECEPCION );
              D.Transaction_Microsip.RollbackRetaining;
              Exit;
            end;
        end;

        D.REPOSITORIO_CFDI.Active := False;
      end
    else // EN CASO CONTRARIO, SOLO HAY QUE HACER LA UNION
      begin
        CFDI_ID := D.IBQueryMicrosip.FieldByName('CFDI_ID').AsInteger;
        XML := D.IBQueryMicrosip.FieldByName('XML').AsString;
      end;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al buscar el repositorio de la compra ' + FOLIO_COMPRA );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end; }

  // COLOCA EL CFDI_RECIBIDO
  { try
    D.GEN_DOCTO_ID.Prepare;
    D.GEN_DOCTO_ID.ExecProc;

    D.CFD_RECIBIDOS.Active := True;
    D.CFD_RECIBIDOS.Insert;
    D.CFD_RECIBIDOS.FieldByName('CFD_RECIBIDO_ID').AsInteger := D.GEN_DOCTO_ID.Params[0].AsInteger;
    D.CFD_RECIBIDOS.FieldByName('CLAVE_SISTEMA').AsString := 'CM';
    D.CFD_RECIBIDOS.FieldByName('DOCTO_ID').AsInteger := DOCTO_CM_ID;
    D.CFD_RECIBIDOS.FieldByName('FECHA').AsDateTime := FECHA_FACTURA;
    D.CFD_RECIBIDOS.FieldByName('XML').AsString := XML;
    D.CFD_RECIBIDOS.FieldByName('CFDI_ID').AsInteger := CFDI_ID;

    D.CFD_RECIBIDOS.Post;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el CFDI recibido ' + FOLIO_COMPRA );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;

  D.CFD_RECIBIDOS.Active := False; }
  {$ENDREGION}



  {$REGION 'GENERAMOS EL CARGO DE LA COMPRA RECIEN CREADA'}
  try
    D.GENERA_DOCTO_CP_CM.ParamByName('V_DOCTO_CM_ID').Value := DOCTO_CM_ID;
    D.GENERA_DOCTO_CP_CM.Prepare;
    D.GENERA_DOCTO_CP_CM.ExecProc;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el cargo de la compra ' + FOLIO_COMPRA );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'OBTENEMOS EL ID DEL CARGO CREADO'}
  try
    D.SELECT.Active := False;
    D.SELECT.SQL.Clear;
    D.SELECT.SQL.Add('SELECT * FROM doctos_entre_sis WHERE docto_fte_id = ' + IntToStr( DOCTO_CM_ID ));
    D.SELECT.Active := True;
    D.SELECT.First;
    DOCTO_CP_ID := D.SELECT.FieldByName('DOCTO_DEST_ID').AsInteger;

    D.SELECT.Active := False;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo cargar el ID del documento en cuentas por pagar ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'COLOCA LOS VENCIMIENTOS DE LA COMPRA ( VENCIMIENTOS_CARGOS_CM )'}
  try
    D.VENCIMIENTOS_CARGOS_CM.TableName := 'VENCIMIENTOS_CARGOS_CM';
    D.VENCIMIENTOS_CARGOS_CM.Active := True;
    D.VENCIMIENTOS_CARGOS_CM.Insert;

    D.VENCIMIENTOS_CARGOS_CM.FieldByName('DOCTO_CM_ID').AsInteger := DOCTO_CM_ID;
    D.VENCIMIENTOS_CARGOS_CM.FieldByName('FECHA_VENCIMIENTO').AsDateTime := FECHA_PAGO;
    D.VENCIMIENTOS_CARGOS_CM.FieldByName('PCTJE_VEN').AsFloat := 100;

    D.VENCIMIENTOS_CARGOS_CM.Post;
    D.VENCIMIENTOS_CARGOS_CM.Active := False;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el vencimiento del documento en compras ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'COLOCA LOS VENCMIENTOS DEL PASIVO ( VENCIMIENTOS_CARGOS_CP )'}
  try
    D.VENCIMIENTOS_CARGOS_CP.TableName := 'VENCIMIENTOS_CARGOS_CP';
    D.VENCIMIENTOS_CARGOS_CP.Active := True;
    D.VENCIMIENTOS_CARGOS_CP.Insert;

    D.VENCIMIENTOS_CARGOS_CP.FieldByName('DOCTO_CP_ID').AsInteger := DOCTO_CP_ID;
    D.VENCIMIENTOS_CARGOS_CP.FieldByName('FECHA_VENCIMIENTO').AsDateTime := FECHA_PAGO;
    D.VENCIMIENTOS_CARGOS_CP.FieldByName('PCTJE_VEN').AsFloat := 100;

    D.VENCIMIENTOS_CARGOS_CP.Post;
    D.VENCIMIENTOS_CARGOS_CP.Active := False;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el vencimiento del documento en cuenta por pagar ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'ACTUALIZO EL ESTATUS DE LA RECEPCIÓN EN MICROSIP'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    D.IBQueryMicrosip.SQL.Add('UPDATE DOCTOS_CM SET ESTATUS = ''F'' WHERE FOLIO = ''' + FOLIO_RECEPCION + ''' AND TIPO_DOCTO = ''R''');
    D.IBQueryMicrosip.Active := True;
  except
    on E : Exception do
      begin
        EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo actualizar el estatus de la recepcion ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'ACTUALIZA EL ESTATUS DE LA FACTURA Y LA RECEPCIÓN EN EL PORTAL'}
  // D.Transaction_Microsip.Commit;
  // D.Transaction_Microsip.Rollback;

  if ( ACTUALIZAR_FACTURA_PORTAL( FOLIO_FINAL, FOLIO_RECEPCION, IntToStr( DOCTO_CM_ID ), RECEP_ID ) = True ) then
    begin
      D.Transaction_Microsip.Commit;

      // SI ESTA CONFIGURADO PARA ENVIO DEL CORREO AUTOMATICO, LO ENVIA
      if ( D.MAILS_SEND = 'True' ) then
        begin
          PROCESO_ENVIAR( PROVEEDOR_ID, RzEditFolioFac, RzEditFecha_fac, RzEditPecha_prov );
        end;
    end
  else
    begin
      D.Transaction_Microsip.Rollback;

      EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), 'Proceso interrumpido', '', '' );
    end; // }
  {$ENDREGION}
end;
{$ENDREGION}

Function SELECT_FACTURAS_APLICAR():Boolean;
  var
    DOCTO_CM_ID, FOLIO_COMPRA, RECEPCION_ID, RECEP_ID, FOLIO_RECEPCION, UUID, MONEDA_SIMBOLO, RFC, NOMBRE, EmpresaID, EmpresaN :String;
    IMPORTE_NETO, TOTAL_IMPUESTOS, TOTAL_RETENCIONES, DESCUENTO_GLOBAL, TIPO_CAMBIO :Double;
    FECHA_PAGO, FECHA_FACTURA, FECHA_RECEPCION, FECHA :TDateTime;
    PROVEEDOR_ID :Integer;
    Fmt :TFormatSettings;

    List :TStringList;
    Line :string;
begin
  List := TStringList.Create;

  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat :='hh:nn:ss';
  Fmt.TimeSeparator  :=':';

  if ( FileExists( ExtractFilePath( ParamStr( 0 ) ) + '/Update/Facturas' ) ) then
    begin
      try
        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Factura.Close;
        D.JvCsvDataSet_Factura.FileName := ExtractFilePath( ParamStr( 0 ) ) + '/Update/Facturas';
        D.JvCsvDataSet_Factura.Open;

        D.JvCsvDataSet_Factura.First;
        while not D.JvCsvDataSet_Factura.Eof do
          begin
            Application.ProcessMessages;

            // CONEXIÓN MICROSIP ( SI ES QUE CAMBIA DE EMPRESA )
            if ( D.Conexion_Microsip.DatabaseName <> ( D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Factura.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB' ) ) then
              begin
                D.Conexion_Microsip.Connected := False;
                D.Conexion_Microsip.DatabaseName := D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Factura.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB';
                D.Conexion_Microsip.Connected := True;
                D.Transaction_Microsip.Active := True;
              end;

            DOCTO_CM_ID := D.JvCsvDataSet_Factura.FieldByName('DOCTO_CM_ID').AsString;
            FOLIO_COMPRA := D.JvCsvDataSet_Factura.FieldByName('FOLIO_COMPRA').AsString;
            IMPORTE_NETO := D.JvCsvDataSet_Factura.FieldByName('IMPORTE_NETO').AsFloat;
            TOTAL_IMPUESTOS := D.JvCsvDataSet_Factura.FieldByName('TOTAL_IMPUESTOS').AsFloat;
            TOTAL_RETENCIONES := D.JvCsvDataSet_Factura.FieldByName('TOTAL_RETENCIONES').AsFloat;
            DESCUENTO_GLOBAL := D.JvCsvDataSet_Factura.FieldByName('DESCUENTO_GLOBAL').AsFloat;
            MONEDA_SIMBOLO := D.JvCsvDataSet_Factura.FieldByName('MONEDA_SIMBOLO').AsString;
            TIPO_CAMBIO := D.JvCsvDataSet_Factura.FieldByName('TIPO_CAMBIO').AsFloat;
            RECEPCION_ID := D.JvCsvDataSet_Factura.FieldByName('RECEPCION_ID').AsString;
            RECEP_ID := D.JvCsvDataSet_Factura.FieldByName('RECEP_ID').AsString;
            FOLIO_RECEPCION := D.JvCsvDataSet_Factura.FieldByName('FOLIO_RECEPCION').AsString;
            FECHA_PAGO := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA_PAGO').AsString, Fmt ); // FECHA_PAGO := D.JvCsvDataSet_FA.FieldByName('FECHA_PAGO').AsDateTime;
            FECHA_FACTURA := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA_FACTURA').AsString, Fmt ); // FECHA_FACTURA := D.JvCsvDataSet_FA.FieldByName('FECHA_FACTURA').AsDateTime;
            FECHA_RECEPCION := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA_RECEPCION').AsString, Fmt ); // FECHA_RECEPCION := D.JvCsvDataSet_FA.FieldByName('FECHA_RECEPCION').AsDateTime;
            FECHA := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA').AsString, Fmt ); // FECHA := D.JvCsvDataSet_FA.FieldByName('FECHA').AsDateTime;
            PROVEEDOR_ID := D.JvCsvDataSet_Factura.FieldByName('PROVEEDOR_ID').AsInteger;
            RFC := D.JvCsvDataSet_Factura.FieldByName('RFC').AsString;
            NOMBRE := D.JvCsvDataSet_Factura.FieldByName('NOMBRE').AsString;
            UUID := D.JvCsvDataSet_Factura.FieldByName('UUID').AsString;
            EmpresaID := D.JvCsvDataSet_Factura.FieldByName('EMPRESA_ID').AsString;
            EmpresaN := D.JvCsvDataSet_Factura.FieldByName('EMPRESA_NOMBRE').AsString;

            EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), 'Subiendo facturas a Microsip Folio: ' + FOLIO_COMPRA, '', '' );
            Sleep( 200 );

            if ( FileExists( ExtractFilePath( ParamStr( 0 ) ) + '/Update/XML/' + UUID + '.xml' ) ) then
              begin
                APLICAR_MICROSIP( DOCTO_CM_ID, RECEP_ID, RECEPCION_ID, EmpresaID, FOLIO_RECEPCION, FOLIO_COMPRA, UUID, RFC, NOMBRE, MONEDA_SIMBOLO, PROVEEDOR_ID, FECHA_PAGO, FECHA_FACTURA, FECHA_RECEPCION, FECHA, IMPORTE_NETO, TOTAL_IMPUESTOS, TOTAL_RETENCIONES, DESCUENTO_GLOBAL, TIPO_CAMBIO, Line );

                if ( Line <> '' ) then
                  begin
                    List.Clear;
                    List.Add('"UUID","FOLIO_XML","FECHA_FACTURA","RFC","NOMBRE","IMPORTE","MONEDA_SIMBOLO","TIPO_CAMBIO","NOMBRE_ARCH","FOLIO_COMPRA","DOCTO_CM_ID","EMPRESA_ID","EMPRESA_NOMBRE"');
                    List.Add( Line + ',"' + EmpresaID + '","' + EmpresaN + '"' );
                    List.SaveToFile( ExtractFilePath( ParamStr( 0 ) ) + '/Update/XML/' + UUID );
                  end;

                // DeleteFile( PChar( ExtractFilePath( ParamStr( 0 ) ) + '/Update/XML/' + UUID + '.xml' ) );
                // D.XML_FILE.Active := False;
              end;

            Inc( D.Position );
            EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '' );

            D.JvCsvDataSet_Factura.Next;
          end;

        // NOS DESCONECTAMOS
        D.Transaction_Microsip.Active := False;
        D.Conexion_Microsip.Connected := False;

        Result := True;
        D.JvCsvDataSet_Factura.Close;
      except
        on E : Exception do
          begin
            EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al cargar las facturas por aplicar' );
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile( PChar( ExtractFilePath( ParamStr( 0 ) ) + '/Update/Facturas' ) );
    end
  else
    begin
      Result := True;
    end;

  List.Destroy;
end;

end.
