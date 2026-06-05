unit Func_Complementos;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Winapi.ShellAPI, WinSvc,
  DateUtils, IdBaseComponent, IdComponent, IdTCPConnection, IdTCPClient, IdMessageClient, IdSMTP, IdMessage,
  XMLDoc, xmldom, XMLIntf,
  // System.Net.HttpClient, System.Net.URLClient,   // THTTPClient
  ComObj, Variants,                 // WinHTTP por COM (TLS lo maneja Windows)
  System.JSON,                                    // parseo de la lista
  System.Zip;                                     // TZipFile
  // System.Classes, System.SysUtils, Data.DB;       // TBytes, TMemoryStream, ftBlob;

  // INSERCIÓN DE FACTURAS 3.3
  Function SELECT_COMPLEMENTOS_APLICAR():Boolean;

implementation

uses
  Data, Func;

const
  // API_BASE = 'http://localhost:8081';   // ej. http://192.168.x.x:8081  (sin / al final)
  API_BASE = 'https://lightgray-bear-551026.hostingersite.com/PortalProveedores/public';   // ej. http://192.168.x.x:8081  (sin / al final)
  API_KEY  = 'acf2c85a4b64f1c21731ee4a94e82e6de0af4cf03b04c39c';        // está en PortalProveedores/.env -> portal.apiKey

type
  TAdjuntoInfo = record
    Id: Integer;
    NombreOriginal: string;
    Tamano: Integer;
  end;

function VariantBytesToTBytes(const V: OleVariant): TBytes;
var
  Lo, Hi, Size: Integer;
  P: Pointer;
begin
  SetLength(Result, 0);
  if VarIsArray(V) then
  begin
    Lo := VarArrayLowBound(V, 1);
    Hi := VarArrayHighBound(V, 1);
    Size := Hi - Lo + 1;
    if Size > 0 then
    begin
      SetLength(Result, Size);
      P := VarArrayLock(V);
      try
        Move(P^, Result[0], Size);
      finally
        VarArrayUnlock(V);
      end;
    end;
  end;
end;

// --- Lista los adjuntos de un documento desde la API del portal ---
function LISTAR_ADJUNTOS_PORTAL(const DoctoId, Emp, Tipo: string): TArray<TAdjuntoInfo>;
var
  Req: OleVariant;
  Bytes: TBytes;
  RespStr: string;
  Raiz: TJSONObject;
  Arr: TJSONArray;
  Item: TJSONObject;
  I: Integer;
begin
  SetLength(Result, 0);
  Req := CreateOleObject('WinHttp.WinHttpRequest.5.1');
  // tiempos (ms): resolución, conexión, envío, recepción
  Req.SetTimeouts(30000, 30000, 30000, 120000);
  Req.Open('GET', Format('%s/api/adjuntos?docto_id=%s&emp=%s&tipo=%s',
                         [API_BASE, DoctoId, Emp, Tipo]), False);
  Req.SetRequestHeader('X-API-Key', API_KEY);
  // Req.Option(9) := 2048;     // <- fuerza TLS 1.2 (solo si corres en Windows 7)
  // Req.Option(4) := 13056;    // <- ignora errores de certificado (solo si es self-signed)
  Req.Send;

  if Req.Status <> 200 then Exit;

  // Decodificamos el cuerpo como UTF-8 a mano (acentos correctos)
  Bytes := VariantBytesToTBytes(Req.ResponseBody);
  RespStr := TEncoding.UTF8.GetString(Bytes);

  Raiz := TJSONObject.ParseJSONValue(RespStr) as TJSONObject;
  if Raiz = nil then Exit;
  try
    Arr := Raiz.GetValue('adjuntos') as TJSONArray;
    if Assigned(Arr) then
    begin
      SetLength(Result, Arr.Count);
      for I := 0 to Arr.Count - 1 do
      begin
        Item := Arr.Items[I] as TJSONObject;
        Result[I].Id             := (Item.GetValue('id') as TJSONNumber).AsInt;
        Result[I].NombreOriginal := Item.GetValue('nombre_original').Value;
        Result[I].Tamano         := (Item.GetValue('tamano') as TJSONNumber).AsInt;
      end;
    end;
  finally
    Raiz.Free;
  end;
end;

// --- Descarga el binario de un adjunto por id ---
function DESCARGAR_ADJUNTO(const Id: Integer; out Contenido: TBytes): Boolean;
var
  Req: OleVariant;
begin
  Result := False;
  SetLength(Contenido, 0);
  Req := CreateOleObject('WinHttp.WinHttpRequest.5.1');
  Req.SetTimeouts(30000, 30000, 30000, 120000);
  Req.Open('GET', Format('%s/api/adjuntos/%d', [API_BASE, Id]), False);
  Req.SetRequestHeader('X-API-Key', API_KEY);
  // Req.Option(9) := 2048;     // TLS 1.2 en Windows 7
  // Req.Option(4) := 13056;    // ignorar cert self-signed
  Req.Send;

  if Req.Status = 200 then
  begin
    Contenido := VariantBytesToTBytes(Req.ResponseBody);  // binario crudo
    Result := True;
  end;
end;

// --- Comprime el contenido en un ZIP de una sola entrada (lo que exige Microsip) ---
function COMPRIMIR_EN_ZIP(const NombreArchivo: string; const Contenido: TBytes): TBytes;
var
  Zip: TZipFile;
  MS: TMemoryStream;
begin
  SetLength(Result, 0);
  MS := TMemoryStream.Create;
  try
    Zip := TZipFile.Create;
    try
      Zip.Open(MS, zmWrite);
      Zip.Add(Contenido, NombreArchivo, zcDeflate);   // entrada nombrada IGUAL que el archivo
      Zip.Close;
    finally
      Zip.Free;
    end;
    SetLength(Result, MS.Size);
    if MS.Size > 0 then
    begin
      MS.Position := 0;
      MS.ReadBuffer(Result[0], MS.Size);
    end;
  finally
    MS.Free;
  end;
end;

// --- Inserta un adjunto en ARCHIVOS_ADJUNTOS de Microsip (dentro de la transacción activa) ---
procedure INSERTAR_ADJUNTO_MICROSIP(const ELEM_DOCTO_CM_ID: Integer;
  const NombreArchivo: string; const TamanoBytes: Integer; const ZipBytes: TBytes);
var
  MS: TMemoryStream;
  ARCHIVO_ADJUNTO_ID: Integer;
begin
  // Nuevo ID con el mismo generador que usa el resto de la función
  D.GEN_DOCTO_ID.Prepare;
  D.GEN_DOCTO_ID.ExecProc;
  ARCHIVO_ADJUNTO_ID := D.GEN_DOCTO_ID.Params[0].AsInteger;

  D.ARCHIVOS_ADJUNTOS_Q.SQL.Clear;
  D.ARCHIVOS_ADJUNTOS_Q.SQL.Add('INSERT INTO ARCHIVOS_ADJUNTOS (');
  D.ARCHIVOS_ADJUNTOS_Q.SQL.Add('  ARCHIVO_ADJUNTO_ID, NOM_TABLA, ELEM_ID, FILE_NAME, FILE_SIZE, FILE_DATE, FILE_STREAM');
  D.ARCHIVOS_ADJUNTOS_Q.SQL.Add(') VALUES (');
  D.ARCHIVOS_ADJUNTOS_Q.SQL.Add('  :ARCHIVO_ADJUNTO_ID, :NOM_TABLA, :ELEM_ID, :FILE_NAME, :FILE_SIZE, :FILE_DATE, :FILE_STREAM');
  D.ARCHIVOS_ADJUNTOS_Q.SQL.Add(')');

  D.ARCHIVOS_ADJUNTOS_Q.ParamByName('ARCHIVO_ADJUNTO_ID').AsInteger := ARCHIVO_ADJUNTO_ID;
  D.ARCHIVOS_ADJUNTOS_Q.ParamByName('NOM_TABLA').AsString          := 'DOCTOS_CM';
  D.ARCHIVOS_ADJUNTOS_Q.ParamByName('ELEM_ID').AsInteger           := ELEM_DOCTO_CM_ID;
  D.ARCHIVOS_ADJUNTOS_Q.ParamByName('FILE_NAME').AsString          := Copy(NombreArchivo, 1, 100);
  D.ARCHIVOS_ADJUNTOS_Q.ParamByName('FILE_SIZE').AsInteger         := TamanoBytes div 1024;  // KB
  D.ARCHIVOS_ADJUNTOS_Q.ParamByName('FILE_DATE').AsDateTime        := Now;

  MS := TMemoryStream.Create;
  try
    if Length(ZipBytes) > 0 then
      MS.WriteBuffer(ZipBytes[0], Length(ZipBytes));
    MS.Position := 0;
    D.ARCHIVOS_ADJUNTOS_Q.ParamByName('FILE_STREAM').LoadFromStream(MS, ftBlob);
  finally
    MS.Free;
  end;

  D.ARCHIVOS_ADJUNTOS_Q.ExecSQL;
end;

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








{$REGION 'ACTUALIZAR_FACTURA_PORTAL_33 - FUNCIÓN QUE ACTUALIZA LAS RECEPCIONES Y FACTURAS EN MYSQL'}
Function ACTUALIZAR_FACTURA_PORTAL_33(FOLIO_COMPRA, FOLIO_RECEPCION, DOCTO_CM_ID, RECEP_ID :string):Boolean;
  var
    CadenaSQL :string;
begin
  // CAMBIO LOS ESTATUS EN LAS FACTURAS EN EL PORTAL A RECIBIDA
  try
    CadenaSQL := CadenaSQL + 'UPDATE COMPLEMENTO_ENCABEZADO SET ';
    CadenaSQL := CadenaSQL + '       ESTATUS = ''R'', ';
    CadenaSQL := CadenaSQL + '       USUARIO_ASOCIO_COBRO = ''' + 'SYSDBA' + ''', ';
    CadenaSQL := CadenaSQL + '       FECHA_ASOCIO_COBRO = ''' + FormatDateTime( 'YYYY-MM-DD hh:nn:ss', Now ) + '''';
    CadenaSQL := CadenaSQL + ' WHERE CREDITO_FK = ' + RECEP_ID;

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
    D.MySQL_Command.CommandText := 'UPDATE CREDITOS SET ESTATUS = ''R'' WHERE CREDITO_FK = ' + RECEP_ID;
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




{$REGION 'APLICAR_MICROSIP_33'}
procedure APLICAR_MICROSIP_33(DOCTO_CM_ID_MYSQL, RECEP_ID, RECEPCION_ID, EMPRESA_ID, FOLIO_RECEPCION, FOLIO_COMPRA, UUID, RFC, NOMBRE, MONEDA_SIMBOLO :String; PROVEEDOR_ID :Integer; FECHA_PAGO, FECHA_FACTURA, FECHA_RECEPCION, FECHA :TDateTime; IMPORTE_NETO, TOTAL_IMPUESTOS, TOTAL_RETENCIONES, DESCUENTO_GLOBAL, TIPO_CAMBIO :Double);
  var
    CadenaSQL, RzEditFolioFac, RzEditFecha_fac, RzEditPecha_prov, FOLIO_FINAL, FOLIO_XML, XML, LUGAR_EXPEDICION, USO_CFDI :String;
    XMLA :AnsiString;
    DOCTO_CM_ID, DOCTO_CM_DET_ID, DOCTO_CM_LIGA_ID, CFDI_ID, DOCTO_CP_ID, IMPTE_DOCTO_CP_ID :Integer;
    XML_FILE :TXMLDocument;

    Utf8Bytes: TBytes;
    Latin1Encoding: TEncoding;

    ListaAdj: TArray<TAdjuntoInfo>; Contenido, ZipBytes: TBytes; I: Integer;
begin
  {$REGION 'BUSCO EL ID DEL CREDITO Y DEL PROVEEDOR EN MICROSIP POR FOLIO Y PROVEEDOR'}
  try
    D.IBQueryMicrosip.Active := False;
    D.IBQueryMicrosip.SQL.Clear;
    D.IBQueryMicrosip.SQL.Add('SELECT');
    D.IBQueryMicrosip.SQL.Add('       DOCTO_CP_ID, ');
    D.IBQueryMicrosip.SQL.Add('       CLAVE_PROV, ');
    D.IBQueryMicrosip.SQL.Add('       PROVEEDOR_ID, ');
    D.IBQueryMicrosip.SQL.Add('       TIPO_CAMBIO, ');
    D.IBQueryMicrosip.SQL.Add('       DESCRIPCION, ');
    D.IBQueryMicrosip.SQL.Add('       COND_PAGO_ID, ');
    D.IBQueryMicrosip.SQL.Add('       CONCEPTO_CP_ID, ');
    D.IBQueryMicrosip.SQL.Add('       TIENE_CFD ');
    D.IBQueryMicrosip.SQL.Add('  FROM DOCTOS_CP');
    D.IBQueryMicrosip.SQL.Add(' WHERE FOLIO = ''' + FOLIO_RECEPCION + '''');
    D.IBQueryMicrosip.SQL.Add('   AND PROVEEDOR_ID = ' + IntToStr(PROVEEDOR_ID));
    D.IBQueryMicrosip.Active := True;
    D.IBQueryMicrosip.Last;

    if (D.IBQueryMicrosip.RecordCount = 0) then
      begin
        {$REGION 'SI NO HUBO RENGLONES ENTONCES NO EXISTE LA RECEPCIÓN Y SE SALE DEL PROCESO'}
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'No se puede asociar el CFDI porque no exise el pago ' + FOLIO_RECEPCION);
        Exit;
        {$ENDREGION}
      end;

    DOCTO_CM_ID := D.IBQueryMicrosip.FIELDBYNAME('DOCTO_CP_ID').AsInteger;
    TIPO_CAMBIO := D.IBQueryMicrosip.FieldByName('TIPO_CAMBIO').AsFloat;

    if (D.IBQueryMicrosip.FieldByName('TIENE_CFD').AsString = 'S') then
      begin
        {$REGION 'SI LA RECEPCION ESTA FACTURADA SOLO ACTUALIZA EL PORTAL Y SE SALE'}
        try
          { D.SELECT.Active := False;
          D.SELECT.SQL.Clear;
          D.SELECT.SQL.Add('SELECT d.*, dc.FOLIO FROM doctos_cm_ligas d');
          D.SELECT.SQL.Add('  JOIN doctos_cm dc ON ( d.docto_cm_dest_id = dc.docto_cm_id )');
          D.SELECT.SQL.Add(' WHERE d.docto_cm_fte_id = ' + IntToStr(DOCTO_CM_ID));
          D.SELECT.Active := True;

          // OBTENEMOS EL FOLIO DE LA COMPRA YA CAPTURADA
          FOLIO_COMPRA := D.SELECT.FieldByName('FOLIO').AsString; }

          // EN ESTA PARTE USA FOLIO_COMPRA EN VEZ DE FOLIO_FINAL PORQUE ES UN FOLIO QUE YA ESTA REGISTRADO
          ACTUALIZAR_FACTURA_PORTAL_33(FOLIO_COMPRA, FOLIO_RECEPCION, IntToStr(DOCTO_CM_ID), RECEP_ID);

          Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', 'No se puede asociar el CFDI porque el credito ' + FOLIO_RECEPCION + ' ya tiene un CFDI asociado');
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
          // D.ADOQueryMySQL.SQL.Add('       F.LUGAR_EXPEDICION, ');
          D.ADOQueryMySQL.SQL.Add('       F.USO_CFDI ');
          D.ADOQueryMySQL.SQL.Add('  FROM COMPLEMENTO_ARCHIVO A ');
          D.ADOQueryMySQL.SQL.Add('  JOIN COMPLEMENTO_ENCABEZADO F ON (A.DOCTO_CP_FK = F.DOCTO_CP_ID)');
          D.ADOQueryMySQL.SQL.Add(' WHERE A.UUID = ''' + UUID + '''');

          D.ADOQueryMySQL.Active := True;

          // LUGAR_EXPEDICION := D.ADOQueryMySQl.FieldByName('LUGAR_EXPEDICION').AsString;
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


end;
{$ENDREGION}









Function SELECT_COMPLEMENTOS_APLICAR():Boolean;
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

  if (FileExists(ExtractFilePath(ParamStr(0)) + '/Update/Complementos')) then
    begin
      try
        D.Conexion_MySQL.Connected := False;
        D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
        D.Conexion_MySQL.Connected := True;

        D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
        D.MySQL_Command.Execute;

        D.JvCsvDataSet_Factura.Close;
        D.JvCsvDataSet_Factura.FileName := ExtractFilePath( ParamStr( 0 ) ) + '/Update/Complementos';
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

            DOCTO_CM_ID := D.JvCsvDataSet_Factura.FieldByName('DOCTO_CP_ID').AsString;
            FOLIO_COMPRA := D.JvCsvDataSet_Factura.FieldByName('FOLIO_PAGO').AsString;
            IMPORTE_NETO := D.JvCsvDataSet_Factura.FieldByName('MONTO').AsFloat;
            // TOTAL_IMPUESTOS := D.JvCsvDataSet_Factura.FieldByName('TOTAL_IMPUESTOS').AsFloat;
            // TOTAL_RETENCIONES := D.JvCsvDataSet_Factura.FieldByName('TOTAL_RETENCIONES').AsFloat;
            // DESCUENTO_GLOBAL := D.JvCsvDataSet_Factura.FieldByName('DESCUENTO_GLOBAL').AsFloat;
            MONEDA_SIMBOLO := D.JvCsvDataSet_Factura.FieldByName('MONEDA_PAGO').AsString;
            // TIPO_CAMBIO := D.JvCsvDataSet_Factura.FieldByName('TIPO_CAMBIO').AsFloat;
            // RECEPCION_ID := D.JvCsvDataSet_Factura.FieldByName('RECEPCION_ID').AsString;
            RECEP_ID := D.JvCsvDataSet_Factura.FieldByName('CREDITO_FK').AsString;
            FOLIO_RECEPCION := D.JvCsvDataSet_Factura.FieldByName('FOLIO_CREDITO').AsString;
            FECHA_PAGO := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA_PAGO').AsString, Fmt ); // FECHA_PAGO := D.JvCsvDataSet_FA.FieldByName('FECHA_PAGO').AsDateTime;
            FECHA_FACTURA := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA_COMPLEMENTO').AsString, Fmt ); // FECHA_FACTURA := D.JvCsvDataSet_FA.FieldByName('FECHA_FACTURA').AsDateTime;
            // FECHA_RECEPCION := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA_RECEPCION').AsString, Fmt ); // FECHA_RECEPCION := D.JvCsvDataSet_FA.FieldByName('FECHA_RECEPCION').AsDateTime;
            // FECHA := StrToDateTime( D.JvCsvDataSet_Factura.FieldByName('FECHA').AsString, Fmt ); // FECHA := D.JvCsvDataSet_FA.FieldByName('FECHA').AsDateTime;
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
      // DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + '/Update/Complementos'));
    end
  else
    begin
      Result := True;
    end;
end;





end.
