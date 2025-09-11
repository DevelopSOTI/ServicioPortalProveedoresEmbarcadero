unit Func_Facturas_3_3;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf;

  // INSERCIÓN DE FACTURAS 3.3
  Function SELECT_FACTURAS_APLICAR_33():Boolean;

implementation

uses
  Data, Func;


{$REGION 'ACTUALIZAR_FACTURA_PORTAL_33 - FUNCIÓN QUE ACTUALIZA LAS RECEPCIONES Y FACTURAS EN MYSQL'}
Function ACTUALIZAR_FACTURA_PORTAL_33(FOLIO_COMPRA, FOLIO_RECEPCION, DOCTO_CM_ID, RECEP_ID :string):Boolean;
  var
    CadenaSQL :string;
begin
  // CAMBIO LOS ESTATUS EN LAS FACTURAS EN EL PORTAL A RECIBIDA
  try
    CadenaSQL := CadenaSQL + 'UPDATE FACTURA_PROVEEDOR_33 SET ';
    CadenaSQL := CadenaSQL + '       FOLIO_MSP = ''' + FOLIO_COMPRA + ''', ';
    CadenaSQL := CadenaSQL + '       COMPRA_ID = ' + DOCTO_CM_ID + ', ';
    CadenaSQL := CadenaSQL + '       ESTATUS = ''R'', ';
    CadenaSQL := CadenaSQL + '       USUARIO_CONV_COMPRA = ''' + 'SYSDBA' + ''', ';
    CadenaSQL := CadenaSQL + '       FECHA_CONV_COMPRA = ''' + FormatDateTime( 'YYYY-MM-DD hh:nn:ss', Now ) + '''';
    CadenaSQL := CadenaSQL + ' WHERE RECEP_ID = ' + RECEP_ID;

    D.MySQL_Command.CommandText := CadenaSQL;
    D.MySQL_Command.Execute;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo actualizar el estatus de la factura ' + FOLIO_COMPRA + ' en el portal');
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
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo actualizar el estatus de la recepcion ' + FOLIO_RECEPCION + ' en el portal web');
        D.Transaction_Microsip.RollbackRetaining;
        Result := False;
        Exit;
      end;
  end;

  Result := True;
end;
{$ENDREGION}

{ procedure LoadXMLFromString(const XMLText: string; XMLDoc: IXMLDocument);
var
  CleanedXML: string;
begin
  CleanedXML := XMLText;

  // Eliminar el BOM si existe (como caracteres ï»¿ o su valor Unicode #$FEFF)
  if (Length(CleanedXML) > 0) and (CleanedXML[1] = #$FEFF) then
    Delete(CleanedXML, 1, 1);

  XMLDoc.LoadFromXML(CleanedXML);
  XMLDoc.Active := True;
end; }

function CleanXMLText(const XMLText: string): string;
begin
  Result := XMLText;

  // Caso 1: BOM como carácter Unicode real (U+FEFF)
  if (Length(Result) > 0) and (Result[1] = #$FEFF) then
    Delete(Result, 1, 1);

  // Caso 2: BOM convertido a caracteres visibles "ï»¿"
  if Result.StartsWith('ï»¿') then
    Delete(Result, 1, 3);
end;

// procedure LoadXMLFromString(const XMLText: string; XMLDoc: IXMLDocument);
{ Function LoadXMLFromString(const XMLText: string; XMLDoc: TXMLDocument): TXMLDocument;
var
  CleanedXML: string;
begin
  CleanedXML := CleanXMLText(XMLText);
  XMLDoc.LoadFromXML(CleanedXML);
  XMLDoc.Active := True;

  Result := XMLDoc;
end; }

{function SafeISO88591String(const UnicodeStr: string): AnsiString;
var
  Encoding: TEncoding;
begin
  Encoding := TEncoding.GetEncoding(28591); // ISO-8859-1
  Result := Encoding.GetBytes(UnicodeStr);  // Esto puede fallar si hay caracteres incompatibles
end;  }

{ function SafeISO88591String(const UnicodeStr: string): AnsiString;
var
  Encoding: TEncoding;
  Bytes: TBytes;
begin
  Encoding := TEncoding.GetEncoding(28591); // ISO-8859-1
  Bytes := Encoding.GetBytes(UnicodeStr);
  SetString(Result, PAnsiChar(@Bytes[0]), Length(Bytes)); // Convierte TBytes a AnsiString
end; }

{ function CodificarISO8859_1(CadenaOriginal :string): AnsiString;
var
  // CadenaOriginal: string;
  Bytes: TBytes;
  CadenaCodificada: AnsiString;
  EncodingISO: TEncoding;
begin
  // CadenaOriginal := 'Texto con acentos: áéíóú ñ';

  // Obtener codificación ISO-8859-1 (Latin1)
  EncodingISO := TEncoding.GetEncoding(28591); // 28591 = ISO-8859-1

  // Convertir la cadena a bytes usando esa codificación
  Bytes := EncodingISO.GetBytes(CadenaOriginal);

  // Convertir los bytes a una AnsiString (1 byte por carácter)
  SetLength(CadenaCodificada, Length(Bytes));
  Move(Bytes[0], CadenaCodificada[1], Length(Bytes));

  // Mostrar resultado
  // Writeln('Cadena codificada (ISO-8859-1): ', CadenaCodificada);
  Result := CadenaCodificada;
end; }

{ function BytesToHex(const ABytes: TBytes): string;
var
  I: Integer;
begin
  Result := '';
  for I := 0 to Length(ABytes) - 1 do
    Result := Result + IntToHex(ABytes[I], 2); // Convierte cada byte a hexadecimal de 2 dígitos
end; }












{$REGION 'APLICAR_MICROSIP_33'}
procedure APLICAR_MICROSIP_33(DOCTO_CM_ID_MYSQL, RECEP_ID, RECEPCION_ID, EMPRESA_ID, FOLIO_RECEPCION, FOLIO_COMPRA, UUID, RFC, NOMBRE, MONEDA_SIMBOLO :String; PROVEEDOR_ID :Integer; FECHA_PAGO, FECHA_FACTURA, FECHA_RECEPCION, FECHA :TDateTime; IMPORTE_NETO, TOTAL_IMPUESTOS, TOTAL_RETENCIONES, DESCUENTO_GLOBAL, TIPO_CAMBIO :Double);
  var
    CadenaSQL, RzEditFolioFac, RzEditFecha_fac, RzEditPecha_prov, FOLIO_FINAL, FOLIO_XML, XML, LUGAR_EXPEDICION, USO_CFDI :String;
    XMLA :AnsiString;
    DOCTO_CM_ID, DOCTO_CM_DET_ID, DOCTO_CM_LIGA_ID, CFDI_ID, DOCTO_CP_ID, IMPTE_DOCTO_CP_ID :Integer;
    XML_FILE :TXMLDocument;

    Utf8Bytes: TBytes;
    Latin1Encoding: TEncoding;
begin
  {$REGION 'BUSCO EL ID DE LA RECEPCIÓN Y DEL PROVEEDOR EN MICROSIP POR FOLIO Y PROVEEDOR'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    // D.IBQueryMicrosip.SQL.Add('SELECT * FROM DOCTOS_CM');
    D.IBQueryMicrosip.SQL.Add('SELECT');
    D.IBQueryMicrosip.SQL.Add('       DOCTO_CM_ID, ');
    D.IBQueryMicrosip.SQL.Add('       ESTATUS, ');
    D.IBQueryMicrosip.SQL.Add('       SUBTIPO_DOCTO, ');
    D.IBQueryMicrosip.SQL.Add('       SUCURSAL_ID, ');
    D.IBQueryMicrosip.SQL.Add('       CLAVE_PROV, ');
    D.IBQueryMicrosip.SQL.Add('       PROVEEDOR_ID, ');
    D.IBQueryMicrosip.SQL.Add('       ALMACEN_ID, ');
    D.IBQueryMicrosip.SQL.Add('       MONEDA_ID, ');
    D.IBQueryMicrosip.SQL.Add('       TIPO_CAMBIO, ');
    D.IBQueryMicrosip.SQL.Add('       TIPO_DSCTO, ');
    D.IBQueryMicrosip.SQL.Add('       DSCTO_PCTJE, ');
    D.IBQueryMicrosip.SQL.Add('       DSCTO_IMPORTE, ');
    D.IBQueryMicrosip.SQL.Add('       DESCRIPCION, ');
    D.IBQueryMicrosip.SQL.Add('       IMPORTE_NETO, ');
    D.IBQueryMicrosip.SQL.Add('       FLETES, ');
    D.IBQueryMicrosip.SQL.Add('       OTROS_CARGOS, ');
    D.IBQueryMicrosip.SQL.Add('       TOTAL_IMPUESTOS, ');
    D.IBQueryMicrosip.SQL.Add('       TOTAL_RETENCIONES, ');
    D.IBQueryMicrosip.SQL.Add('       GASTOS_ADUANALES, ');
    D.IBQueryMicrosip.SQL.Add('       OTROS_GASTOS, ');
    D.IBQueryMicrosip.SQL.Add('       COND_PAGO_ID, ');
    D.IBQueryMicrosip.SQL.Add('       CARGAR_SUN, ');
    D.IBQueryMicrosip.SQL.Add('       CONSIG_CM_ID, ');
    D.IBQueryMicrosip.SQL.Add('       PEDIMENTO_ID, ');
    D.IBQueryMicrosip.SQL.Add('       VIA_EMBARQUE_ID, ');
    D.IBQueryMicrosip.SQL.Add('       IMPUESTO_SUSTITUIDO_ID, ');
    D.IBQueryMicrosip.SQL.Add('       IMPUESTO_SUSTITUTO_ID ');
    D.IBQueryMicrosip.SQL.Add('  FROM DOCTOS_CM');
    D.IBQueryMicrosip.SQL.Add(' WHERE FOLIO = ''' + FOLIO_RECEPCION + '''');
    D.IBQueryMicrosip.SQL.Add('   AND PROVEEDOR_ID = ' + IntToStr(PROVEEDOR_ID));
    D.IBQueryMicrosip.SQL.Add('   AND TIPO_DOCTO = ''R''');
    D.IBQueryMicrosip.Active := True;
    D.IBQueryMicrosip.Last;

    if (D.IBQueryMicrosip.RecordCount = 0) then
      begin
        {$REGION 'SI NO HUBO RENGLONES ENTONCES NO EXISTE LA RECEPCIÓN Y SE SALE DEL PROCESO'}
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'No se puede insertar la compra porque no exise la recepción ' + FOLIO_RECEPCION);
        Exit;
        {$ENDREGION}
      end;

    DOCTO_CM_ID := D.IBQueryMicrosip.FIELDBYNAME('DOCTO_CM_ID').AsInteger;

    if (D.IBQueryMicrosip.FieldByName('ESTATUS').AsString = 'F') then
      begin
        {$REGION 'SI LA RECEPCION ESTA FACTURADA SOLO ACTUALIZA EL PORTAL Y SE SALE'}
        try
          D.SELECT.Active := False;
          D.SELECT.SQL.Clear;
          D.SELECT.SQL.Add('SELECT d.*, dc.FOLIO FROM doctos_cm_ligas d');
          D.SELECT.SQL.Add('  JOIN doctos_cm dc ON ( d.docto_cm_dest_id = dc.docto_cm_id )');
          D.SELECT.SQL.Add(' WHERE d.docto_cm_fte_id = ' + IntToStr(DOCTO_CM_ID));
          D.SELECT.Active := True;

          // OBTENEMOS EL FOLIO DE LA COMPRA YA CAPTURADA
          FOLIO_COMPRA := D.SELECT.FieldByName('FOLIO').AsString;

          // EN ESTA PARTE USA FOLIO_COMPRA EN VEZ DE FOLIO_FINAL PORQUE ES UN FOLIO QUE YA ESTA REGISTRADO
          ACTUALIZAR_FACTURA_PORTAL_33(FOLIO_COMPRA, FOLIO_RECEPCION, IntToStr(DOCTO_CM_ID), RECEP_ID);

          Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'No se puede insertar la compra porque la recepción ' + FOLIO_RECEPCION + ' ya fue facturada');
          Exit;
        except
          on E : Exception do
            begin
              Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al intentar cargar el identificador de la compra de la recepción ' + FOLIO_RECEPCION);
              Exit;
            end;
        end;
        {$ENDREGION}
      end;
  except
    on E:Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al cargar los datos de Microsip (Recepción)');
        Exit;
      end;
  end;
  {$ENDREGION}

  if (FOLIO_COMPRA <> '000000000') then
    begin
      {$REGION 'SI LA FACTURA TIENE FOLIO DIFERENTE DE 0 BUSCA QUE NO ESTE REGISTRADA CON EL MISMO PROVEEDOR'}
      try
        D.SELECT.Active := False;
        D.SELECT.SQL.Clear;
        D.SELECT.SQL.Add('SELECT DOCTO_CM_ID FROM DOCTOS_CM' );
        D.SELECT.SQL.Add(' WHERE FOLIO_PROV = ''' + FOLIO_COMPRA + '''');
        D.SELECT.SQL.Add('   AND PROVEEDOR_ID = ' + IntToStr( PROVEEDOR_ID ));
        D.SELECT.SQL.Add('   AND TIPO_DOCTO = ''C''');
        D.SELECT.Active := True;
        D.SELECT.Last;

        if ( D.SELECT.RecordCount <> 0 ) then
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'No se puede insertar porque el folio de la factura ' + FOLIO_COMPRA + ' ya se encuentra registrado');
            Exit;
          end;
      except
        on E:Exception do
          begin
            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al revisar si la factura ' + FOLIO_COMPRA + ' ya se encontraba registrada');
            Exit;
          end;
      end;
      {$ENDREGION}
    end;

  // SI PASO TODAS LAS VALIDACIONES, ENTONCES PROCEDE A OBTENER EL NUEVO FOLIO WEB
  FOLIO_FINAL := SIGUIENTE_FOLIO('WEB');

  if (FOLIO_FINAL = 'WEB000000') then
    begin
      EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'La serie WEB no esta registrada en la empresa ' + NOMBRE);
      Exit;
    end;

  {$REGION 'SI VIENE EN 0 EL FOLIO DE LA COMPRA QUIERE DECIR QUE EL XML NO TENIA FOLIO Y ASIGNA COMO FOLIO DE PROVEEDOR EL FOLIO WEB'}
  if (FOLIO_COMPRA <> '000000000') then
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
  RzEditFecha_fac := FormatDateTime('dd/mm/yyyy', FECHA_FACTURA);
  RzEditPecha_prov := FormatDateTime('dd/mm/yyyy', FECHA_RECEPCION);

  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando compra ' + IntToStr(DOCTO_CM_ID));

  {$REGION 'COLOCA EL ENCABEZADO DE LA COMPRA (DOCTOS_CM)'}
  try
    D.GEN_DOCTO_ID.Prepare;
    D.GEN_DOCTO_ID.ExecProc;
    DOCTO_CM_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

    D.DOCTOS_CM_Q.SQL.Clear;
    D.DOCTOS_CM_Q.SQL.Add('INSERT INTO DOCTOS_CM (');
    D.DOCTOS_CM_Q.SQL.Add('  DOCTO_CM_ID, TIPO_DOCTO, SUBTIPO_DOCTO, SUCURSAL_ID, FOLIO, FECHA, CLAVE_PROV,');
    D.DOCTOS_CM_Q.SQL.Add('  PROVEEDOR_ID, FOLIO_PROV, FACTURA_DEV, ALMACEN_ID, MONEDA_ID, TIPO_CAMBIO,');
    D.DOCTOS_CM_Q.SQL.Add('  TIPO_DSCTO, DSCTO_PCTJE, DSCTO_IMPORTE, ESTATUS, APLICADO, DESCRIPCION,');
    D.DOCTOS_CM_Q.SQL.Add('  IMPORTE_NETO, FLETES, OTROS_CARGOS, TOTAL_IMPUESTOS, TOTAL_RETENCIONES,');
    D.DOCTOS_CM_Q.SQL.Add('  GASTOS_ADUANALES, OTROS_GASTOS, FORMA_EMITIDA, CONTABILIZADO,');
    D.DOCTOS_CM_Q.SQL.Add('  ACREDITAR_CXP, SISTEMA_ORIGEN, COND_PAGO_ID, PCTJE_DSCTO_PPAG, CARGAR_SUN, ENVIADO,');
    D.DOCTOS_CM_Q.SQL.Add('  TIENE_CFD, USUARIO_CREADOR, USUARIO_AUT_CREACION, USUARIO_ULT_MODIF, USUARIO_AUT_MODIF');
    // D.DOCTOS_CM_Q.SQL.Add('  CONSIG_CM_ID, PEDIMENTO_ID, VIA_EMBARQUE_ID, IMPUESTO_SUSTITUIDO_ID, IMPUESTO_SUSTITUTO_ID');
    D.DOCTOS_CM_Q.SQL.Add(') VALUES (');
    D.DOCTOS_CM_Q.SQL.Add('  :DOCTO_CM_ID, :TIPO_DOCTO, :SUBTIPO_DOCTO, :SUCURSAL_ID, :FOLIO, :FECHA, :CLAVE_PROV,');
    D.DOCTOS_CM_Q.SQL.Add('  :PROVEEDOR_ID, :FOLIO_PROV, :FACTURA_DEV, :ALMACEN_ID, :MONEDA_ID, :TIPO_CAMBIO,');
    D.DOCTOS_CM_Q.SQL.Add('  :TIPO_DSCTO, :DSCTO_PCTJE, :DSCTO_IMPORTE, :ESTATUS, :APLICADO, :DESCRIPCION,');
    D.DOCTOS_CM_Q.SQL.Add('  :IMPORTE_NETO, :FLETES, :OTROS_CARGOS, :TOTAL_IMPUESTOS, :TOTAL_RETENCIONES,');
    D.DOCTOS_CM_Q.SQL.Add('  :GASTOS_ADUANALES, :OTROS_GASTOS, :FORMA_EMITIDA, :CONTABILIZADO,');
    D.DOCTOS_CM_Q.SQL.Add('  :ACREDITAR_CXP, :SISTEMA_ORIGEN, :COND_PAGO_ID, :PCTJE_DSCTO_PPAG, :CARGAR_SUN, :ENVIADO,');
    D.DOCTOS_CM_Q.SQL.Add('  :TIENE_CFD, :USUARIO_CREADOR, :USUARIO_AUT_CREACION, :USUARIO_ULT_MODIF, :USUARIO_AUT_MODIF');
    // D.DOCTOS_CM_Q.SQL.Add('  :CONSIG_CM_ID, :PEDIMENTO_ID, :VIA_EMBARQUE_ID, :IMPUESTO_SUSTITUIDO_ID, :IMPUESTO_SUSTITUTO_ID');
    D.DOCTOS_CM_Q.SQL.Add(')');

    with D.DOCTOS_CM_Q.Params do
    begin
      ParamByName('DOCTO_CM_ID').AsInteger := DOCTO_CM_ID;
      ParamByName('TIPO_DOCTO').AsString := 'C';
      ParamByName('SUBTIPO_DOCTO').AsString := D.IBQueryMicrosip.FieldByName('SUBTIPO_DOCTO').AsString;
      ParamByName('SUCURSAL_ID').AsInteger := D.IBQueryMicrosip.FieldByName('SUCURSAL_ID').AsInteger;
      ParamByName('FOLIO').AsString := FOLIO_FINAL;
      ParamByName('FECHA').AsDateTime := FECHA;
      ParamByName('CLAVE_PROV').AsString := D.IBQueryMicrosip.FieldByName('CLAVE_PROV').AsString;
      ParamByName('PROVEEDOR_ID').AsInteger := D.IBQueryMicrosip.FieldByName('PROVEEDOR_ID').AsInteger;
      ParamByName('FOLIO_PROV').AsString := FOLIO_COMPRA;
      ParamByName('FACTURA_DEV').AsString := '';
      ParamByName('ALMACEN_ID').AsInteger := D.IBQueryMicrosip.FieldByName('ALMACEN_ID').AsInteger;
      ParamByName('MONEDA_ID').AsInteger := D.IBQueryMicrosip.FieldByName('MONEDA_ID').AsInteger;
      ParamByName('TIPO_CAMBIO').AsFloat := D.IBQueryMicrosip.FieldByName('TIPO_CAMBIO').AsFloat;
      ParamByName('TIPO_DSCTO').AsString := D.IBQueryMicrosip.FieldByName('TIPO_DSCTO').AsString;
      ParamByName('DSCTO_PCTJE').AsFloat := D.IBQueryMicrosip.FieldByName('DSCTO_PCTJE').AsFloat;
      ParamByName('DSCTO_IMPORTE').AsFloat := D.IBQueryMicrosip.FieldByName('DSCTO_IMPORTE').AsFloat;
      ParamByName('ESTATUS').AsString := 'N';
      ParamByName('APLICADO').AsString := 'S';
      ParamByName('DESCRIPCION').AsString := D.IBQueryMicrosip.FieldByName('DESCRIPCION').AsString;
      ParamByName('IMPORTE_NETO').AsFloat := D.IBQueryMicrosip.FieldByName('IMPORTE_NETO').AsFloat;
      ParamByName('FLETES').AsFloat := D.IBQueryMicrosip.FieldByName('FLETES').AsFloat;
      ParamByName('OTROS_CARGOS').AsFloat := D.IBQueryMicrosip.FieldByName('OTROS_CARGOS').AsFloat;
      ParamByName('TOTAL_IMPUESTOS').AsFloat := D.IBQueryMicrosip.FieldByName('TOTAL_IMPUESTOS').AsFloat;
      ParamByName('TOTAL_RETENCIONES').AsFloat := D.IBQueryMicrosip.FieldByName('TOTAL_RETENCIONES').AsFloat;
      ParamByName('GASTOS_ADUANALES').AsFloat := D.IBQueryMicrosip.FieldByName('GASTOS_ADUANALES').AsFloat;
      ParamByName('OTROS_GASTOS').AsFloat := D.IBQueryMicrosip.FieldByName('OTROS_GASTOS').AsFloat;
      ParamByName('FORMA_EMITIDA').AsString := 'N';
      ParamByName('CONTABILIZADO').AsString := 'N';
      ParamByName('ACREDITAR_CXP').AsString := 'N';
      ParamByName('SISTEMA_ORIGEN').AsString := 'CM';
      ParamByName('COND_PAGO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('COND_PAGO_ID').AsInteger;
      ParamByName('PCTJE_DSCTO_PPAG').AsFloat := 0;
      ParamByName('CARGAR_SUN').AsString := D.IBQueryMicrosip.FieldByName('CARGAR_SUN').AsString;
      ParamByName('ENVIADO').AsString := 'N';
      ParamByName('TIENE_CFD').AsString := 'S';
      ParamByName('USUARIO_CREADOR').AsString := 'SISTEMAWEB';
      ParamByName('USUARIO_AUT_CREACION').AsString := 'SYSDBA';
      ParamByName('USUARIO_ULT_MODIF').AsString := 'SISTEMAWEB';
      ParamByName('USUARIO_AUT_MODIF').AsString := 'SISTEMAWEB';

      // Opcionales, usar solo si son distintos de 0
      { ParamByName('CONSIG_CM_ID').AsInteger := D.IBQueryMicrosip.FieldByName('CONSIG_CM_ID').AsInteger;
      ParamByName('PEDIMENTO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('PEDIMENTO_ID').AsInteger;
      ParamByName('VIA_EMBARQUE_ID').AsInteger := D.IBQueryMicrosip.FieldByName('VIA_EMBARQUE_ID').AsInteger;
      ParamByName('IMPUESTO_SUSTITUIDO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('IMPUESTO_SUSTITUIDO_ID').AsInteger;
      ParamByName('IMPUESTO_SUSTITUTO_ID').AsInteger := D.IBQueryMicrosip.FieldByName('IMPUESTO_SUSTITUTO_ID').AsInteger; // }
    end;

    D.DOCTOS_CM_Q.ExecSQL;

    {$REGION 'COLOCA EL ENCABEZADO DE LA COMPRA (DOCTOS_CM)'}

    { D.GEN_DOCTO_ID.Prepare;
    D.GEN_DOCTO_ID.ExecProc;
    DOCTO_CM_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

    // D.DOCTOS_CM_Q.SQL.Text := '';
    // D.DOCTOS_CM_Q.RequestLive := True;
    // D.DOCTOS_CM_Q.
    // D.DOCTOS_CM_Q.Open;

    D.DOCTOS_CM.TableName := 'DOCTOS_CM';
    D.DOCTOS_CM.Active := True;
    D.DOCTOS_CM.Insert;

    D.DOCTOS_CM.FieldByName('DOCTO_CM_ID').AsInteger := DOCTO_CM_ID;
    D.DOCTOS_CM.FieldByName('TIPO_DOCTO').AsString := 'C';
    D.DOCTOS_CM.FieldByName('SUBTIPO_DOCTO').AsString := D.IBQueryMicrosip.FieldByName('SUBTIPO_DOCTO').AsString;
    D.DOCTOS_CM.FieldByName('SUCURSAL_ID').AsInteger := D.IBQueryMicrosip.FieldByName('SUCURSAL_ID').AsInteger;
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
    // D.DOCTOS_CM.FieldByName('TIENE_CFD').AsString := 'N';
    D.DOCTOS_CM.FieldByName('TIENE_CFD').AsString := 'S';
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
    D.DOCTOS_CM.Active := False; // }

    {$ENDREGION}
  except
    on E : Exception do
      begin
        Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el encabezado del documento de compra ' + FOLIO_COMPRA );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando ligas ' + IntToStr(DOCTO_CM_ID));

  {$REGION 'COLOCA EL DOCTOS_CM_LIGAS (DOCTOS_CM_LIGAS) (ENTRE RECEPCIÓN Y LA COMPRA CREADA)'}
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
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el encabezado de la liga de la rececpción ' + FOLIO_RECEPCION + ' a la compra ' + FOLIO_COMPRA);
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando detalle ' + IntToStr(DOCTO_CM_ID));

  {$REGION 'COLOCA EL DETALLE DE LA COMPRA Y LAS LIGAS CORRESPONDIENTES (DOCTOS_CM_DET Y DOCTOS_CM_LIGAS_DET)'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;

    D.IBQueryMicrosip.SQL.Add('SELECT D.* FROM DOCTOS_CM_DET D');
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
              Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el detalle del documento de compra ' + FOLIO_COMPRA );
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
              Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el detalle de la liga de la recepción ' + FOLIO_RECEPCION + ' con la compra ' + FOLIO_COMPRA );
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
        Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo cargar el detalle de la rececpción ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando impuestos ' + IntToStr(DOCTO_CM_ID));

  {$REGION 'COLOCA LOS IMPUESTOS DE LA COMPRA (IMPUESTOS_DOCTOS_CM)'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    D.IBQueryMicrosip.SQL.Add( 'SELECT D.* FROM IMPUESTOS_DOCTOS_CM D' );
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
              Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el detalle de impuestos del documento de compra ' + FOLIO_COMPRA);
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
        Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo cargar el detalle de impuestos de la recepción ' + FOLIO_RECEPCION );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}





  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando repositorio ' + IntToStr(DOCTO_CM_ID));

  {$REGION 'COLOCA/BUSCA EL REPOSITORIO_CFDI - CARGA LOS ARCHIVOS QUE ESTAN EN EL PORTAL DE LA FACTURA EN PROCESO'}
  try
    // BUSCA SI ESTA EL REPOSITORIO_CFDI
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    D.IBQueryMicrosip.SQL.Add( 'SELECT CFDI_ID, XML FROM REPOSITORIO_CFDI WHERE UUID = ''' + UUID + '''');
    D.IBQueryMicrosip.Active := True;
    D.IBQueryMicrosip.Last;

    // SI NO ESTA EL REPOSITORIO HAY QUE CREARLO
    if ( D.IBQueryMicrosip.RecordCount = 0 ) then
      begin
        try
          D.ADOQueryMySQL.Active := False;
          D.ADOQueryMySQL.SQL.Clear;
          // D.ADOQueryMySQL.SQL.Add('SELECT XML, LUGAR_EXPEDICION, USO_CFDI FROM ARCHIVOS_FACTURA_PROVEEDOR_33');
          // D.ADOQueryMySQL.SQL.Add(' WHERE UUID = ''' + UUID + '''');

          D.ADOQueryMySQL.SQL.Add('SELECT ');
          // D.ADOQueryMySQL.SQL.Add('       A.XML, ');
          D.ADOQueryMySQL.SQL.Add('       CONVERT(XML USING utf8) AS XML, ');
          D.ADOQueryMySQL.SQL.Add('       F.LUGAR_EXPEDICION, ');
          D.ADOQueryMySQL.SQL.Add('       F.USO_CFDI ');
          D.ADOQueryMySQL.SQL.Add('  FROM ARCHIVOS_FACTURA_PROVEEDOR_33 A ');
          D.ADOQueryMySQL.SQL.Add('  JOIN FACTURA_PROVEEDOR_33 F ON (A.DOCTO_CM_FK = F.DOCTO_CM_ID)');
          D.ADOQueryMySQL.SQL.Add(' WHERE A.UUID = ''' + UUID + '''');

          D.ADOQueryMySQL.Active := True;

          LUGAR_EXPEDICION := D.ADOQueryMySQl.FieldByName('LUGAR_EXPEDICION').AsString;
          USO_CFDI := D.ADOQueryMySQl.FieldByName('USO_CFDI').AsString;
          XML := D.ADOQueryMySQl.FieldByName('XML').AsString;

          // XML_FILE := TXMLDocument.Create(nil);
          // XMLDoc := TXMLDocument.Create(nil);
          // XML_FILE := LoadXMLFromString(XML, XML_FILE);
          { // XML_FILE.Active := False;
          // XML_FILE.Options := [doNodeAutoIndent];
          XML_FILE.Active := True;
          XML_FILE.Version := '1.0';
          XML_FILE.LoadFromXML(XML);
          XML_FILE.Encoding := 'ISO-8859-1'; }



          XML := CleanXMLText(XML);
          // XML := UTF8Encode(XML);
          // XMLA := CodificarISO8859_1(XML);

          // XMLA := SafeISO88591String(XML);
          // XMLA := SafeISO88591String(XML);

          // XML_FILE.LoadFromXML(XML);
          // XML_FILE.Encoding := 'ISO-8859-1';
          // XML_FILE.Active := False;
          // XML_FILE.Options := [doNodeAutoIndent];
          // XML_FILE.Encoding := 'UTF-8';



          // XML := xmlDoc.FormatXMLData(XML_FILE.XML.Text);

          // XML := xmlDoc.FormatXMLData(XML_FILE.XML.Text);
          // XML := StringReplace(XML, '<?xml version="1.0"?>', '<?xml version="1.0" encoding="UTF-8"?>', [rfReplaceAll]); }




          // Utf8Bytes := TEncoding.UTF8.GetBytes(XML);

          // Puedes convertir los bytes de nuevo a cadena si lo necesitas
          // Writeln('Cadena original: ', OriginalString);
          // Writeln('Cadena en UTF-8 (en bytes): ', BytesToHex(Utf8Bytes));


          // XML := TEncoding.UTF8.GetString(TEncoding.Convert(TEncoding.UTF8, TEncoding.GetEncoding(28591), TEncoding.UTF8.GetBytes(XML)));
          Utf8Bytes := TEncoding.Convert(TEncoding.UTF8, TEncoding.GetEncoding(28591), TEncoding.UTF8.GetBytes(XML));
          XML := TEncoding.ASCII.GetString(Utf8Bytes);
          // XML := TEncoding.ut.GetString(Utf8Bytes);
          // Latin1Encoding := TEncoding.GetEncoding(28591);

          // Latin1Encoding := TEncoding.GetEncoding(28591, TEncoderFallback.ReplacementFallback, TDecoderFallback.ReplacementFallback);
          // XML := TEncoding.UTF8.GetString(TEncoding.Convert(TEncoding.UTF8, Latin1Encoding, TEncoding.UTF8.GetBytes(XML)));


          // ShowMessage(XML);
          // ShowMessage(BytesToHex(Utf8Bytes));
          // ShowMessage(XMLA);

          D.GEN_DOCTO_ID.Prepare;
          D.GEN_DOCTO_ID.ExecProc;
          CFDI_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

          D.REPOSITORIO_CFDI_Q.SQL.Clear;
          D.REPOSITORIO_CFDI_Q.SQL.Add('INSERT INTO REPOSITORIO_CFDI (');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  CFDI_ID, MODALIDAD_FACTURACION, VERSION, UUID, NATURALEZA, TIPO_COMPROBANTE,');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  TIPO_DOCTO_MSP, FOLIO, FECHA, RFC, NOMBRE, IMPORTE, MONEDA, TIPO_CAMBIO,');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  ES_PARCIALIDAD, NOM_ARCH, XML, REFER_GRUPO, SELLO_VALIDADO, ES_SUSTITUTO,');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  USUARIO_CREADOR, FECHA_HORA_CREACION, LUGAR_EXPEDICION, USO_CFDI');
          D.REPOSITORIO_CFDI_Q.SQL.Add(') VALUES (');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  :CFDI_ID, :MODALIDAD_FACTURACION, :VERSION, :UUID, :NATURALEZA, :TIPO_COMPROBANTE,');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  :TIPO_DOCTO_MSP, :FOLIO, :FECHA, :RFC, :NOMBRE, :IMPORTE, :MONEDA, :TIPO_CAMBIO,');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  :ES_PARCIALIDAD, :NOM_ARCH, :XML, :REFER_GRUPO, :SELLO_VALIDADO, :ES_SUSTITUTO,');
          D.REPOSITORIO_CFDI_Q.SQL.Add('  :USUARIO_CREADOR, :FECHA_HORA_CREACION, :LUGAR_EXPEDICION, :USO_CFDI');
          D.REPOSITORIO_CFDI_Q.SQL.Add(')');

          with D.REPOSITORIO_CFDI_Q.Params do
          begin
            ParamByName('CFDI_ID').AsInteger := CFDI_ID;
            ParamByName('MODALIDAD_FACTURACION').AsString := 'CFDI';
            ParamByName('VERSION').AsString := '4.0';
            ParamByName('UUID').AsString := UUID;
            ParamByName('NATURALEZA').AsString := 'R';
            ParamByName('TIPO_COMPROBANTE').AsString := 'I';
            ParamByName('TIPO_DOCTO_MSP').AsString := 'Compra';
            ParamByName('FOLIO').AsString := FOLIO_XML;
            ParamByName('FECHA').AsDateTime := FECHA_FACTURA;
            ParamByName('RFC').AsString := RFC;
            ParamByName('NOMBRE').AsString := NOMBRE;
            ParamByName('IMPORTE').AsFloat := IMPORTE_NETO + TOTAL_IMPUESTOS - TOTAL_RETENCIONES - DESCUENTO_GLOBAL;
            ParamByName('MONEDA').AsString := MONEDA_SIMBOLO;
            ParamByName('TIPO_CAMBIO').AsFloat := TIPO_CAMBIO;
            ParamByName('ES_PARCIALIDAD').AsString := 'N';
            ParamByName('NOM_ARCH').AsString := RFC + '_' + FOLIO_COMPRA + '.xml';
            ParamByName('XML').AsString := XML;
            ParamByName('REFER_GRUPO').AsString := FOLIO_COMPRA;
            ParamByName('SELLO_VALIDADO').AsString := 'M';
            ParamByName('ES_SUSTITUTO').AsString := 'N';
            ParamByName('USUARIO_CREADOR').AsString := 'SISTEMAWEB';
            ParamByName('FECHA_HORA_CREACION').AsDateTime := Now;
            ParamByName('LUGAR_EXPEDICION').AsString := LUGAR_EXPEDICION;
            ParamByName('USO_CFDI').AsString := USO_CFDI;
          end;

          D.REPOSITORIO_CFDI_Q.ExecSQL;

          {$REGION 'COLOCA/BUSCA EL REPOSITORIO_CFDI - CARGA LOS ARCHIVOS QUE ESTAN EN EL PORTAL DE LA FACTURA EN PROCESO'}

          { D.GEN_DOCTO_ID.Prepare;
          D.GEN_DOCTO_ID.ExecProc;
          CFDI_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

          D.REPOSITORIO_CFDI.Active := True;
          D.REPOSITORIO_CFDI.Insert;

          D.REPOSITORIO_CFDI.FieldByName('CFDI_ID').AsInteger := CFDI_ID;
          D.REPOSITORIO_CFDI.FieldByName('MODALIDAD_FACTURACION').AsString := 'CFDI';
          D.REPOSITORIO_CFDI.FieldByName('VERSION').AsString := '3.3';
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
          D.REPOSITORIO_CFDI.FieldByName('ES_SUSTITUTO').AsString := 'N';
          D.REPOSITORIO_CFDI.FieldByName('USUARIO_CREADOR').AsString := 'SISTEMAWEB';
          D.REPOSITORIO_CFDI.FieldByName('FECHA_HORA_CREACION').AsDateTime := Now;

          D.REPOSITORIO_CFDI.Post; // }

          {$ENDREGION}
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
        EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al buscar el repositorio de la compra ' + FOLIO_COMPRA);
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;

  // COLOCA EL CFDI_RECIBIDO
  try
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

  D.CFD_RECIBIDOS.Active := False;
  {$ENDREGION}





  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando cargo ' + IntToStr(DOCTO_CM_ID));

  {$REGION 'COLOCA LOS VENCIMIENTOS DE LA COMPRA (VENCIMIENTOS_CARGOS_CM)'}
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
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el vencimiento del documento en compras ' + FOLIO_RECEPCION);
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'GENERAMOS EL CARGO DE LA COMPRA RECIEN CREADA EN CUENTAS POR PAGAR'}
  try
    D.GENERA_DOCTO_CP_CM.ParamByName('V_DOCTO_CM_ID').Value := DOCTO_CM_ID;
    D.GENERA_DOCTO_CP_CM.Prepare;
    D.GENERA_DOCTO_CP_CM.ExecProc;
  except
    on E : Exception do
      begin
        EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el cargo de la compra ' + FOLIO_COMPRA);
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'OBTENEMOS EL ID DEL CARGO CREADO - YA NO SE USA 2025'}
  { try
    D.SELECT.Active := False;
    D.SELECT.SQL.Clear;
    D.SELECT.SQL.Add('SELECT * FROM doctos_entre_sis WHERE docto_fte_id = ' + IntToStr(DOCTO_CM_ID) + ' AND clave_sis_fte = ''CM'' AND clave_sis_dest = ''CP''');
    D.SELECT.Active := True;
    D.SELECT.First;
    DOCTO_CP_ID := D.SELECT.FieldByName('DOCTO_DEST_ID').AsInteger;

    D.SELECT.Active := False;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo cargar el ID del documento en cuentas por pagar ' + FOLIO_RECEPCION);
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end; }
  {$ENDREGION}

  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando cargo ' + IntToStr(DOCTO_CP_ID));

  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando vencimientos cm');

  {$REGION 'COLOCA LOS VENCIMIENTOS DE LA COMPRA (VENCIMIENTOS_CARGOS_CM) - YA NO SE USA 2025'}
  { try
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
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el vencimiento del documento en compras ' + FOLIO_RECEPCION);
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end; }
  {$ENDREGION}

  // EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'Guardando vencimientos cp');

  {$REGION 'COLOCA LOS VENCMIENTOS DEL PASIVO ( VENCIMIENTOS_CARGOS_CP ) - YA NO SE USA 2025'}
  { try
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
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo guardar el vencimiento del documento en cuenta por pagar ' + FOLIO_RECEPCION);
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end; }
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
        Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' No se pudo actualizar el estatus de la recepcion ' + FOLIO_RECEPCION + ' en Microsip.' );
        D.Transaction_Microsip.RollbackRetaining;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'ACTUALIZA EL ESTATUS DE LA FACTURA Y LA RECEPCIÓN EN EL PORTAL'}
  // D.Transaction_Microsip.Commit;
  if (ACTUALIZAR_FACTURA_PORTAL_33(FOLIO_FINAL, FOLIO_RECEPCION, IntToStr(DOCTO_CM_ID), RECEP_ID) = True) then
    begin
      D.Transaction_Microsip.Commit;

      // SI ESTA CONFIGURADO PARA ENVIO DEL CORREO AUTOMATICO, LO ENVIA
      if (D.MAILS_SEND = 'True') then
        begin
          PROCESO_ENVIAR(PROVEEDOR_ID, RzEditFolioFac, RzEditFecha_fac, RzEditPecha_prov);
        end;
    end
  else
    begin
      D.Transaction_Microsip.Rollback;

      EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Proceso interrumpido', '', '');
    end;
  {$ENDREGION}
end;
{$ENDREGION}

Function SELECT_FACTURAS_APLICAR_33():Boolean;
  var
    DOCTO_CM_ID, FOLIO_COMPRA, RECEPCION_ID, RECEP_ID, FOLIO_RECEPCION, UUID, MONEDA_SIMBOLO, RFC, NOMBRE, EmpresaID, EmpresaN :String;
    IMPORTE_NETO, TOTAL_IMPUESTOS, TOTAL_RETENCIONES, DESCUENTO_GLOBAL, TIPO_CAMBIO :Double;
    FECHA_PAGO, FECHA_FACTURA, FECHA_RECEPCION, FECHA :TDateTime;
    PROVEEDOR_ID :Integer;
    Fmt :TFormatSettings;
begin
  Fmt.ShortDateFormat := 'dd/mm/yyyy';
  Fmt.DateSeparator := '/';
  Fmt.LongTimeFormat :='hh:nn:ss';
  Fmt.TimeSeparator  :=':';

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Facturas_33')) then
    begin
      try
        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Factura.Close;
        D.JvCsvDataSet_Factura.FileName := ExtractFilePath( ParamStr( 0 ) ) + '/Update/Facturas_33';
        D.JvCsvDataSet_Factura.Open;

        D.JvCsvDataSet_Factura.First;
        while not D.JvCsvDataSet_Factura.Eof do
          begin
            Application.ProcessMessages;

            // CONEXIÓN MICROSIP ( SI ES QUE CAMBIA DE EMPRESA )
            if (D.Conexion_Microsip.DatabaseName <> (D.MICRO_SERV + ':' + D.MICRO_ROOT + D.JvCsvDataSet_Factura.FieldByName('EMPRESA_NOMBRE').AsString + '.FDB')) then
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

            Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Subiendo facturas a Microsip Folio: ' + FOLIO_COMPRA, '', '');
            Sleep(200);

            APLICAR_MICROSIP_33(DOCTO_CM_ID, RECEP_ID, RECEPCION_ID, EmpresaID, FOLIO_RECEPCION, FOLIO_COMPRA, UUID, RFC, NOMBRE, MONEDA_SIMBOLO, PROVEEDOR_ID, FECHA_PAGO, FECHA_FACTURA, FECHA_RECEPCION, FECHA, IMPORTE_NETO, TOTAL_IMPUESTOS, TOTAL_RETENCIONES, DESCUENTO_GLOBAL, TIPO_CAMBIO);

            Inc(D.Position);
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

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
            EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al cargar las facturas por aplicar');
            Result := False;
          end;
      end;

      D.Conexion_MySQL.Connected := False;
      DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Facturas_33'));
    end
  else
    begin
      Result := True;
    end;
end;

end.
