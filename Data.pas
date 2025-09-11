unit Data;

interface

uses
  System.SysUtils, System.Classes, System.Win.Registry, Winapi.Windows, IBX.IBTable,
  IBX.IBStoredProc, Data.Win.ADODB, Data.DB, IBX.IBCustomDataSet, IBX.IBQuery,
  IBX.IBDatabase, Forms, SvCom_Timer, ActiveX, Dialogs, Xml.xmldom, Xml.XMLIntf,
  Xml.Win.msxmldom, Xml.XMLDoc, JvCsvData;

type
  TD = class(TDataModule)
    Conexion_MySQL: TADOConnection;
    ADOQueryParametros: TADOQuery;
    MySQL_Command: TADOCommand;
    Conexion_Config: TIBDatabase;
    Transaction_Config: TIBTransaction;
    IBQueryConfig: TIBQuery;
    ADOQueryEmpresas: TADOQuery;
    Conexion_Microsip: TIBDatabase;
    Transaction_Microsip: TIBTransaction;
    IBQueryMicrosip: TIBQuery;
    ADOQueryActual: TADOQuery;
    ADOQueryMySQL: TADOQuery;
    IBQueryDetalle: TIBQuery;
    ADOQueryFacturas: TADOQuery;
    GEN_DOCTO_ID: TIBStoredProc;
    CFD_RECIBIDOS: TIBTable;
    REPOSITORIO_CFDI: TIBTable;
    GENERA_DOCTO_CP_CM: TIBStoredProc;
    DOCTOS_CM: TIBTable;
    DOCTOS_CM_LIGAS: TIBTable;
    DOCTOS_CM_DET: TIBTable;
    DOCTOS_CM_LIGAS_DET: TIBTable;
    IMPUESTOS_DOCTOS_CM: TIBTable;
    VENCIMIENTOS_CARGOS_CM: TIBTable;
    VENCIMIENTOS_CARGOS_CP: TIBTable;
    SELECT: TIBQuery;
    JvCsvDataSet_Empresa: TJvCsvDataSet;
    svTimerSync: TsvTimer;
    JvCsvDataSet_Almacen: TJvCsvDataSet;
    JvCsvDataSet_Moneda: TJvCsvDataSet;
    JvCsvDataSet_Proveedor: TJvCsvDataSet;
    JvCsvDataSet_Recepcion: TJvCsvDataSet;
    JvCsvDataSet_Libre: TJvCsvDataSet;
    JvCsvDataSet_Factura: TJvCsvDataSet;
    XML_FILE: TXMLDocument;
    JvCsvDataSet_Credito: TJvCsvDataSet;
    IBQueryXML: TIBQuery;
    DOCTOS_CM_Q: TIBQuery;
    REPOSITORIO_CFDI_Q: TIBQuery;
    procedure DataModuleCreate(Sender: TObject);
    procedure svTimerSyncTimer(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    MYSQL_SERV,
    MYSQL_USER,
    MYSQL_PASS,
    MYSQL_DATA,
    MYSQL_PORT,

    MICRO_SERV,
    MICRO_ROOT,
    MICRO_USER,
    MICRO_PASS,

    MAILS_SEND,
    MODE_APPLI :String;

    AplicaFacturas :Boolean;
    LastUpdate_Text :string;
    LastUpdate_Date :TDateTime;

    ProgressMax, Position :Integer;

    Continue :Boolean;
  end;

var
  D: TD;

implementation

uses
  Main, Form, Func, Func_Calcula;

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{$REGION 'DataModuleCreate - EVENTO CREATE, INICIALIZA LAS VARIABLES DE CONEXIÓN'}
procedure TD.DataModuleCreate(Sender: TObject);
  var
    Reg :TRegistry;
begin
  {$REGION 'INICIALIZA LAS VARIABLES DE REGISTROS'}
  MYSQL_SERV := '';
  MYSQL_USER := '';
  MYSQL_PASS := '';
  MYSQL_DATA := '';
  MYSQL_PORT := '';

  MICRO_SERV := '';
  MICRO_ROOT := '';
  MICRO_USER := '';
  MICRO_PASS := '';

  MAILS_SEND := '';
  MODE_APPLI := '';
  {$ENDREGION}

  {$REGION 'LEEMOS LOS REGISTROS'}
  try
    Reg := TRegistry.Create(KEY_READ or KEY_WOW64_64KEY);
    Reg.RootKey := HKEY_LOCAL_MACHINE;
    if (Reg.KeyExists('SOFTWARE\SOTI\Service Portal')) then
      begin
        if (Reg.OpenKey('SOFTWARE\SOTI\Service Portal', False)) then
          begin
            MYSQL_SERV := Reg.ReadString('MYSQL_SERV');
            MYSQL_USER := Reg.ReadString('MYSQL_USER');
            MYSQL_PASS := Reg.ReadString('MYSQL_PASS');
            MYSQL_DATA := Reg.ReadString('MYSQL_DATA');
            MYSQL_PORT := Reg.ReadString('MYSQL_PORT');

            MICRO_SERV := Reg.ReadString('MICRO_SERV');
            MICRO_ROOT := Reg.ReadString('MICRO_ROOT');
            MICRO_USER := Reg.ReadString('MICRO_USER');
            MICRO_PASS := Reg.ReadString('MICRO_PASS');

            MAILS_SEND := Reg.ReadString('MAILS_SEND');
            MODE_APPLI := Reg.ReadString('MODE_APPLI');
            // MODE_APPLI := 'F';

            Continue := True;
          end
        else
          begin
            Continue := False;
          end;
      end
    else
      begin
        Continue := False;
      end;
    Reg.CloseKey;
    Reg.Free;
  except
    MessageBox(0, 'El replicador no pudo acceder a los registros de Windows.', 'Mensaje de replicador', MB_ICONERROR);
    Continue := False;
  end;
  {$ENDREGION}

  // USUARIO Y CONTRASEÑA PARA EL CONFIG
  D.Conexion_Config.Params.Values['user_name'] := D.MICRO_USER;
  D.Conexion_Config.Params.Values['password'] := D.MICRO_PASS;
  D.Conexion_Microsip.Params.Values['lc_ctype'] := 'ISO8859_1';
  // D.Conexion_Microsip.Params.Values['lc_ctype'] := 'UTF8';

  // USUARIO Y CONTRASEÑA PARA LAS EMPRESAS DE MICROSIP
  D.Conexion_Microsip.Params.Values['user_name'] := D.MICRO_USER;
  D.Conexion_Microsip.Params.Values['password'] := D.MICRO_PASS;
  D.Conexion_Microsip.Params.Values['lc_ctype'] := 'ISO8859_1';
  // D.Conexion_Microsip.Params.Values['lc_ctype'] := 'UTF8';

  // SI MODE_APPLI ES 'F' QUIERE DECIR QUE SE EJECUTARA EN MODO VISUAL
  if (MODE_APPLI = 'F') then
    begin
      Application.CreateForm(TM, M);

      Application.CreateForm(TF, F);

      F.FormStyle := fsStayOnTop;
      F.Panel.Caption := '';
      F.Memo.Lines.Clear;
      F.Update;
      F.ShowModal;
      F.Destroy;
    end;
end;
{$ENDREGION}

{$REGION 'svTimerSyncTimer - EVENTO DEL TIMER'}
procedure TD.svTimerSyncTimer(Sender: TObject);
  var
    RutaFichero :string;
    Interval :Integer;
    Started :TDateTime;
    T :TextFile;
    // num_empresas :Integer;
    Format :TFormatSettings;
begin
  CoInitialize( nil );

  ProgressMax := 1; // INICIALIZA LA PROGRESS BAR CON VALOR DE 1
  Position := 1; // REINICIA EL CONTADOR EN 1
  Started := Now; // REGISTRA LA FECHA Y HORA CON LA QUE EMPEZO LA ACTUALIZACIÓN

  Format.ShortDateFormat := 'dd/mm/yyyy';
  Format.DateSeparator := '/';
  Format.LongTimeFormat := 'hh:nn:ss';
  Format.TimeSeparator := ':';

  {$REGION 'OBTENEMOS Y CALCULAMOS EL INTERVALO DE EJECUCIÓN PARA EL TIMER - Interval'}
  Interval := 1;

  RutaFichero := ExtractFilePath(ParamStr(0)) + 'Timer.ini';
  AssignFile(T, RutaFichero);
  Reset(T);
  while not Eof(T) do
    begin
      Readln(T, Interval);
    end;

  Interval := Interval * 1000;
  {$ENDREGION}

  {$REGION 'SI NO EXISTEN LAS CARPETAS "Service Report, EventLog, Update y XML" LAS CREAMOS'}
  try
    if (not DirectoryExists(ExtractFilePath(ParamStr(0)) + 'Service Report')) then
      begin
        CreateDir(ExtractFilePath(ParamStr(0)) + 'Service Report');
      end;
  except
  end;

  try
    if (not DirectoryExists(ExtractFilePath(ParamStr(0)) + 'EventLog')) then
      begin
        CreateDir(ExtractFilePath(ParamStr(0)) + 'EventLog');
      end;
  except
  end;

  try
    if (not DirectoryExists(ExtractFilePath(ParamStr(0)) + 'Update')) then
      begin
        CreateDir(ExtractFilePath(ParamStr(0)) + 'Update');
      end;
  except
  end;

  try
    if (not DirectoryExists(ExtractFilePath(ParamStr(0)) + 'Update/XML')) then
      begin
        CreateDir(ExtractFilePath(ParamStr(0)) + 'Update/XML');
      end;
  except
  end;
  {$ENDREGION}



  // SI NO HA SIDO CONFIGURADO O HUBO UN ERROR AL LEER LOS REGISTROS EL SERVICIO TERMINA EL PROCESO
  if Continue = False then
    begin
      // SyncService.EventLog.LogMessage( 'El servicio no ha sido configurado, el proceso de sincronización fue interrumpido.' );
      Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', 'El servicio no ha sido configurado, el proceso de sincronización fue interrumpido.' );
      Exit;
    end;



  svTimerSync.Enabled := False; // DETIENE EL TIMER PARA INICIAR LA SINCRONIZACIÓN
  svTimerSync.Interval := Interval; // ACTUALIZAMOS EL INTERVALO DEL TIMER

  Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Iniciando la sincronización con el portal', '', '', True);

  {$REGION 'NOS CONECTAMOS AL PORTAL'}
  try
    D.Conexion_MySQL.Connected := False;
    D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
    D.Conexion_MySQL.Connected := True;

    D.MySQL_Command.CommandText := 'SET SQL_BIG_SELECTS = 1';
    D.MySQL_Command.Execute;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Proceso terminado', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al intentar conectarse al portal');

        svTimerSync.Enabled := True;
        Exit;
      end;
  end;
  {$ENDREGION}

  {$REGION 'OBTENEMOS LA ULTIMA FECHA DE MODIFICACIÓN Y REVISAMOS SI VAMOS A INSERTAR LAS FACTURAS DE FORMA AUTOMATICA O MANUAL'}
  try
    Func.GET_APLICA_FACTURAS(); // VEMOS SI APLICAREMOS EN AUTOMATICO LAS FACTURAS O SERA MANUAL EL PROCESO
    Func.GET_LAST_UPDATE(); // OBTENEMOS LA ULTIMA FECHA DE ACTUALIZACIÓN (MENOS UN DIA)

    if (D.LastUpdate_Text <> '') then
      begin
        D.LastUpdate_Date := StrToDateTime(D.LastUpdate_Text, Format);
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', 'Ultima sincronización el ' + FormatDateTime('dddd dd', D.LastUpdate_Date) + ' de ' + FormatDateTime('mmmm', D.LastUpdate_Date) + ' del ' + FormatDateTime('yyyy', D.LastUpdate_Date) + ' a las ' + FormatDateTime('hh:nn:ss am/pm', D.LastUpdate_Date), '');
      end
    else
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', 'No se ha realizado ninguna sincronización', '');
      end;
  except
    on E : Exception do
      begin
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Proceso terminado', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al intentar conectarse al portal');

        svTimerSync.Enabled := True;
        Exit;
      end;
  end;
  {$ENDREGION}

  try
    if (Func_Calcula.CALCULA_REGISTROS(D.ProgressMax) = True) then
      begin
        // SI AL CALCULAR EL NÚMERO DE REGISTROS A PROCESAR NO HAY ERROR CONTINUA CON EL PROCESO
        Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), '', '', '');

        // SI SE EJECUTARA EN MODO VISUAL INICIALIZA EL MAXIMO DE LA PROGRESSBAR
        if (MODE_APPLI = 'F') then
          begin
            F.ProgressBar.Max := D.ProgressMax;
          end;

        if (Func.INSERT_UPDATE_EMPRESAS = True) then // ACTUALIZA LAS EMPRESAS DE MICROSIP
          begin
            if (Func.INSERT_UPDATE = True) then // ACTUALIZA TODA LA INFORMACIÓN DEL PORTAL
              begin
                // ACTUALIZAMOS EL VALOR DE LA ULTIMA SINCRONIZACIÓN
                try
                  Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), 'Finalizando sincronización', '', '' );

                  D.Conexion_MySQL.Connected := False;
                  D.Conexion_MySQL.ConnectionString := 'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=' + D.MYSQL_USER + ';PORT=' + D.MYSQL_PORT + ';DATABASE=' + D.MYSQL_DATA + ';SERVER=' + D.MYSQL_SERV + ';PASSWORD=' + D.MYSQL_PASS + ';';
                  D.Conexion_MySQL.Connected := True;

                  D.MySQL_Command.CommandText := 'UPDATE PARAMETROS SET PARAM_VALOR = ''' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Started) + ''' WHERE PARAM_CLAVE = ''LAST_UPDATE''';
                  D.MySQL_Command.Execute;

                  D.MySQL_Command.CommandText := 'UPDATE EMPRESAS_MSP SET EMP_ULT_SINC = ''' + FormatDateTime('YYYY-MM-DD HH:MM:SS', Started) + ''' WHERE EMP_ESTATUS = ''Autorizada''';
                  D.MySQL_Command.Execute;

                  D.LastUpdate_Date := Started;
                  Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', 'Ultima sincronización el ' + FormatDateTime( 'dddd dd', D.LastUpdate_Date ) + ' de ' + FormatDateTime( 'mmmm', D.LastUpdate_Date ) + ' del ' + FormatDateTime( 'yyyy', D.LastUpdate_Date ) + ' a las ' + FormatDateTime( 'hh:nn:ss am/pm', D.LastUpdate_Date ), '' );
                except
                  on E : Exception do
                    begin
                      Func.EVENT_LOG( IntToStr( D.ProgressMax ), IntToStr( D.Position ), '', '', '[' + E.ClassName + '] ' + E.Message + ' Hubo un error al intentar conectarse al portal' );
                    end;
                end; // }
              end;
          end;
      end;
  except
    on E:Exception do
      begin
        // NO HACER NADA Y CONTINUAR CON LA SINCRONIZACIÓN
      end;
  end;

  {$REGION 'FINALIZAMOS LA SINCRONIZACIÓN, NOS DESCONECTAMOS DE TODAS LAS BASES DE DATOS'}
  try
    D.Conexion_Config.Connected := False; // SE DESCONECTA DEL CONFIG
    D.Conexion_Microsip.Connected := False; // SE DESCONECTA DE MICROSIP
    D.Conexion_MySQL.Connected := False; // SE DESCONECTA DEL PORTAL

    // Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.Position), 'Proceso terminado', '', '', True, 'F');
    Func.EVENT_LOG(IntToStr(D.ProgressMax), IntToStr(D.ProgressMax), 'Proceso terminado', '', '', True, 'F');

    DeleteFile(PChar(ExtractFilePath(ParamStr(0)) + 'EventLog.ini'));
  except
    // NO HACER NADA Y CONTINUAR CON LA SINCRONIZACIÓN
  end;
  {$ENDREGION}

  svTimerSync.Enabled := True; // INICIA EL TIMER DE NUEVO
end;
{$ENDREGION}

end.
