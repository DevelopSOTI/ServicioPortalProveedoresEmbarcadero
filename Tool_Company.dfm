object TC: TTC
  Left = 0
  Top = 0
  Caption = 'Configurar empresas autorizadas y bloqueadas'
  ClientHeight = 345
  ClientWidth = 857
  Color = clBtnFace
  Constraints.MinHeight = 382
  Constraints.MinWidth = 873
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnCreate = FormCreate
  DesignSize = (
    857
    345)
  PixelsPerInch = 96
  TextHeight = 13
  object Label4: TLabel
    Left = 546
    Top = 55
    Width = 303
    Height = 13
    Anchors = [akTop, akRight]
    Caption = 
      'Valores '#39'N'#39' = No permite diferencias, '#39'S'#39' = Si permite diferenci' +
      'as'
  end
  object Label2: TLabel
    Left = 8
    Top = 8
    Width = 841
    Height = 41
    Anchors = [akLeft, akTop, akRight]
    AutoSize = False
    Caption = 
      'Las empresas que se muestran a continuaci'#243'n son las mismas que e' +
      'stan en Microsip, pero no todas seran visibles en el portal de p' +
      'roveedores, solo las empresas que esten autorizadas seran visibl' +
      'es en el portal. El permitir diferencias significa que las factu' +
      'ras que suban los proveedores podran tener diferencias en los im' +
      'portes, siempre y cuando no pasen del limite especificado en los' +
      ' datos particulares del proveedor indicados en Microsip.'
    Color = clBtnFace
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentColor = False
    ParentFont = False
    Transparent = True
    WordWrap = True
  end
  object DBGridEmpresas: TDBGrid
    Left = 8
    Top = 74
    Width = 841
    Height = 263
    Anchors = [akLeft, akTop, akRight, akBottom]
    DataSource = DataSource_Empresas
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit, dgMultiSelect]
    PopupMenu = PopupMenu1
    ReadOnly = True
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'EMP_ID'
        Visible = False
      end
      item
        Expanded = False
        FieldName = 'EMP_NOMBRE'
        Title.Caption = 'Nombre'
        Width = 155
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'EMP_NOMBRE_LARGO'
        Title.Caption = 'Nombre completo'
        Width = 350
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'EMP_RFC'
        Title.Caption = 'RFC'
        Width = 100
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'EMP_ESTATUS'
        Title.Caption = 'Estado'
        Width = 90
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'EMP_DIFERENCIA'
        Title.Caption = 'Permite diferencias'
        Width = 96
        Visible = True
      end>
  end
  object Conexion_MySQL: TADOConnection
    ConnectionString = 
      'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=soticomm_admon;PORT=330' +
      '6;DATABASE=soticomm_RelPro;SERVER=localhost;'
    LoginPrompt = False
    Mode = cmRead
    Left = 48
    Top = 120
  end
  object Select_Empresas: TADOQuery
    Connection = Conexion_MySQL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM EMPRESAS_MSP'
      'ORDER BY EMP_ID')
    Left = 136
    Top = 120
  end
  object DataSource_Empresas: TDataSource
    DataSet = Select_Empresas
    Left = 240
    Top = 120
  end
  object Command_Update: TADOCommand
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 344
    Top = 120
  end
  object Select_Diferencia: TADOQuery
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 440
    Top = 120
  end
  object PopupMenu1: TPopupMenu
    BiDiMode = bdLeftToRight
    MenuAnimation = [maTopToBottom]
    ParentBiDiMode = False
    Left = 520
    Top = 120
    object Autorizar1: TMenuItem
      Caption = 'Autorizar en la pagina'
      OnClick = Autorizar1Click
    end
    object Bloquear1: TMenuItem
      Caption = 'Bloquear en la pagina'
      OnClick = Bloquear1Click
    end
    object Permitediferencias1: TMenuItem
      Caption = 'Permite diferencias'
      OnClick = Permitediferencias1Click
    end
    object Rechazadiferencias1: TMenuItem
      Caption = 'Rechaza diferencias'
      OnClick = Rechazadiferencias1Click
    end
  end
end
