object D: TD
  OldCreateOrder = False
  OnCreate = DataModuleCreate
  Height = 579
  Width = 917
  object Conexion_MySQL: TADOConnection
    LoginPrompt = False
    Mode = cmRead
    Provider = 'MSDASQL.1'
    Left = 112
    Top = 24
  end
  object ADOQueryParametros: TADOQuery
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 112
    Top = 120
  end
  object MySQL_Command: TADOCommand
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 112
    Top = 72
  end
  object Conexion_Config: TIBDatabase
    LoginPrompt = False
    ServerType = 'IBServer'
    Left = 248
    Top = 24
  end
  object Transaction_Config: TIBTransaction
    DefaultDatabase = Conexion_Config
    Params.Strings = (
      'write'
      'concurrency')
    Left = 248
    Top = 72
  end
  object IBQueryConfig: TIBQuery
    Database = Conexion_Config
    Transaction = Transaction_Config
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 248
    Top = 120
  end
  object ADOQueryEmpresas: TADOQuery
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 112
    Top = 168
  end
  object Conexion_Microsip: TIBDatabase
    LoginPrompt = False
    ServerType = 'IBServer'
    Left = 376
    Top = 24
  end
  object Transaction_Microsip: TIBTransaction
    DefaultDatabase = Conexion_Microsip
    Left = 376
    Top = 72
  end
  object IBQueryMicrosip: TIBQuery
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 376
    Top = 120
  end
  object ADOQueryActual: TADOQuery
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 112
    Top = 216
  end
  object ADOQueryMySQL: TADOQuery
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 112
    Top = 264
  end
  object IBQueryDetalle: TIBQuery
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 376
    Top = 168
  end
  object ADOQueryFacturas: TADOQuery
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 112
    Top = 312
  end
  object GEN_DOCTO_ID: TIBStoredProc
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    StoredProcName = 'GEN_DOCTO_ID'
    Left = 376
    Top = 264
  end
  object CFD_RECIBIDOS: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    TableName = 'CFD_RECIBIDOS'
    UniDirectional = False
    Left = 376
    Top = 312
  end
  object REPOSITORIO_CFDI: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    TableName = 'REPOSITORIO_CFDI'
    UniDirectional = False
    Left = 376
    Top = 360
  end
  object GENERA_DOCTO_CP_CM: TIBStoredProc
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    StoredProcName = 'GENERA_DOCTO_CP_CM'
    Left = 376
    Top = 408
  end
  object DOCTOS_CM: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    UniDirectional = False
    Left = 528
    Top = 24
  end
  object DOCTOS_CM_LIGAS: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    UniDirectional = False
    Left = 528
    Top = 72
  end
  object DOCTOS_CM_DET: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    UniDirectional = False
    Left = 528
    Top = 120
  end
  object DOCTOS_CM_LIGAS_DET: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    UniDirectional = False
    Left = 528
    Top = 168
  end
  object IMPUESTOS_DOCTOS_CM: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    UniDirectional = False
    Left = 528
    Top = 216
  end
  object VENCIMIENTOS_CARGOS_CM: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    UniDirectional = False
    Left = 528
    Top = 264
  end
  object VENCIMIENTOS_CARGOS_CP: TIBTable
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    UniDirectional = False
    Left = 528
    Top = 312
  end
  object SELECT: TIBQuery
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 528
    Top = 360
  end
  object JvCsvDataSet_Empresa: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 24
  end
  object svTimerSync: TsvTimer
    Enabled = False
    OnTimer = svTimerSyncTimer
    Left = 808
    Top = 24
  end
  object JvCsvDataSet_Almacen: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 72
  end
  object JvCsvDataSet_Moneda: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 120
  end
  object JvCsvDataSet_Proveedor: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 168
  end
  object JvCsvDataSet_Recepcion: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 216
  end
  object JvCsvDataSet_Libre: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 312
  end
  object JvCsvDataSet_Factura: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 360
  end
  object XML_FILE: TXMLDocument
    Left = 808
    Top = 72
    DOMVendorDesc = 'MSXML'
  end
  object JvCsvDataSet_Credito: TJvCsvDataSet
    AutoBackupCount = 0
    Left = 688
    Top = 264
  end
  object IBQueryXML: TIBQuery
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 376
    Top = 216
  end
  object DOCTOS_CM_Q: TIBQuery
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 528
    Top = 432
  end
  object REPOSITORIO_CFDI_Q: TIBQuery
    Database = Conexion_Microsip
    Transaction = Transaction_Microsip
    BufferChunks = 1000
    CachedUpdates = False
    ParamCheck = True
    Left = 528
    Top = 488
  end
end
