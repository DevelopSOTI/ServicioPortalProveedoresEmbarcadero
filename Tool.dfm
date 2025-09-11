object T: TT
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Configuraci'#243'n del replicador'
  ClientHeight = 370
  ClientWidth = 753
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu
  OldCreateOrder = False
  Position = poScreenCenter
  OnActivate = FormActivate
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 353
    Height = 137
    Caption = 'Conexi'#243'n a Microsip'
    TabOrder = 0
    object Label2: TLabel
      Left = 16
      Top = 24
      Width = 159
      Height = 13
      Caption = 'Nombre o IP del servidor Microsip'
    end
    object Label3: TLabel
      Left = 50
      Top = 51
      Width = 125
      Height = 13
      Caption = 'Ubicaci'#243'n de las empresas'
    end
    object Label4: TLabel
      Left = 139
      Top = 78
      Width = 36
      Height = 13
      Caption = 'Usuario'
    end
    object Label8: TLabel
      Left = 119
      Top = 105
      Width = 56
      Height = 13
      Caption = 'Contrase'#241'a'
    end
    object MICRO_SERV: TEdit
      Left = 181
      Top = 21
      Width = 159
      Height = 21
      TabOrder = 0
      Text = 'localhost'
    end
    object MICRO_ROOT: TEdit
      Left = 181
      Top = 48
      Width = 159
      Height = 21
      TabOrder = 1
      Text = 'C:\Microsip Datos\'
    end
    object MICRO_USER: TEdit
      Left = 181
      Top = 75
      Width = 100
      Height = 21
      ReadOnly = True
      TabOrder = 2
      Text = 'SYSDBA'
    end
    object MICRO_PASS: TEdit
      Left = 181
      Top = 102
      Width = 100
      Height = 21
      PasswordChar = '*'
      TabOrder = 3
    end
  end
  object GroupBox2: TGroupBox
    Left = 367
    Top = 8
    Width = 378
    Height = 137
    Caption = 'Conexi'#243'n al portal de proveedores'
    TabOrder = 1
    object Label10: TLabel
      Left = 15
      Top = 24
      Width = 88
      Height = 13
      Caption = 'Dominio o servidor'
    end
    object Label11: TLabel
      Left = 35
      Top = 51
      Width = 68
      Height = 13
      Caption = 'Base de datos'
    end
    object Label12: TLabel
      Left = 67
      Top = 78
      Width = 36
      Height = 13
      Caption = 'Usuario'
    end
    object Label14: TLabel
      Left = 47
      Top = 105
      Width = 56
      Height = 13
      Caption = 'Contrase'#241'a'
    end
    object Label5: TLabel
      Left = 274
      Top = 24
      Width = 32
      Height = 13
      Caption = 'Puerto'
    end
    object MYSQL_SERV: TEdit
      Left = 109
      Top = 21
      Width = 159
      Height = 21
      TabOrder = 0
      Text = 'localhost'
    end
    object MYSQL_DATA: TEdit
      Left = 109
      Top = 48
      Width = 159
      Height = 21
      TabOrder = 1
    end
    object MYSQL_USER: TEdit
      Left = 109
      Top = 75
      Width = 100
      Height = 21
      TabOrder = 2
    end
    object MYSQL_PASS: TEdit
      Left = 109
      Top = 102
      Width = 100
      Height = 21
      PasswordChar = '*'
      TabOrder = 3
    end
    object MYSQL_PORT: TEdit
      Left = 312
      Top = 21
      Width = 50
      Height = 21
      TabOrder = 4
      Text = '3306'
    end
  end
  object GroupBox3: TGroupBox
    Left = 8
    Top = 151
    Width = 737
    Height = 74
    Caption = 'Otros parametros'
    TabOrder = 2
    object checkAutomatico: TCheckBox
      Left = 16
      Top = 27
      Width = 686
      Height = 14
      Caption = 
        #191'Desea que el servicio registre las compras de forma automatica?' +
        ' Solo se registraran las facturas que esten relacionadas a una r' +
        'ecepci'#243'n.'
      TabOrder = 0
    end
    object checkEnviaCorreo: TCheckBox
      Left = 16
      Top = 47
      Width = 290
      Height = 17
      Caption = 'Enviar correo a los proveedores al generar las compras.'
      TabOrder = 1
    end
    object checkCierraConfig: TCheckBox
      Left = 312
      Top = 47
      Width = 226
      Height = 17
      Caption = 'Cerrar el replicador al terminar el proceso.'
      TabOrder = 2
      Visible = False
    end
  end
  object GroupBox4: TGroupBox
    Left = 8
    Top = 231
    Width = 737
    Height = 90
    Caption = 'Replicador'
    TabOrder = 3
    object Label13: TLabel
      Left = 16
      Top = 30
      Width = 78
      Height = 13
      Caption = 'Sincronizar cada'
    end
    object lblTimer: TLabel
      Left = 16
      Top = 61
      Width = 46
      Height = 13
      Caption = 'lblTimer'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clRed
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object BTN_Install: TButton
      Left = 248
      Top = 25
      Width = 75
      Height = 25
      Caption = 'Instalar'
      Enabled = False
      TabOrder = 1
      OnClick = BTN_InstallClick
    end
    object BTN_Uninstall: TButton
      Left = 410
      Top = 25
      Width = 75
      Height = 25
      Caption = 'Desinstalar'
      Enabled = False
      TabOrder = 2
      OnClick = BTN_UninstallClick
    end
    object BTN_Start: TButton
      Left = 329
      Top = 25
      Width = 75
      Height = 25
      Caption = 'Iniciar'
      Enabled = False
      TabOrder = 3
      OnClick = BTN_StartClick
    end
    object BTN_Stop: TButton
      Left = 491
      Top = 25
      Width = 75
      Height = 25
      Caption = 'Detener'
      Enabled = False
      TabOrder = 4
      OnClick = BTN_StopClick
    end
    object EDIT_TIME: TEdit
      Left = 100
      Top = 27
      Width = 60
      Height = 21
      Alignment = taRightJustify
      TabOrder = 0
    end
    object comboTime: TComboBox
      Left = 166
      Top = 27
      Width = 76
      Height = 21
      Style = csDropDownList
      ItemIndex = 0
      TabOrder = 5
      Text = 'Segundos'
      Items.Strings = (
        'Segundos'
        'Minutos'
        'Horas')
    end
  end
  object Panel1: TPanel
    Left = 0
    Top = 329
    Width = 753
    Height = 41
    Align = alBottom
    BevelOuter = bvNone
    Color = 3355443
    ParentBackground = False
    TabOrder = 4
    DesignSize = (
      753
      41)
    object BTN_ACCEPT: TButton
      Left = 572
      Top = 8
      Width = 75
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Aceptar'
      TabOrder = 0
      OnClick = BTN_ACCEPTClick
    end
    object BTN_CANCEL: TButton
      Left = 653
      Top = 8
      Width = 75
      Height = 25
      Anchors = [akTop, akRight]
      Caption = 'Cancelar'
      TabOrder = 1
      OnClick = BTN_CANCELClick
    end
  end
  object MainMenu: TMainMenu
    Left = 582
    Top = 257
    object Herramientas1: TMenuItem
      Caption = 'Herramientas'
      object Empresas1: TMenuItem
        Caption = 'Configurar empresas autorizadas y bloqueadas'
        OnClick = Empresas1Click
      end
      object Correo1: TMenuItem
        Caption = 'Configurar correo de la p'#225'gina'
        OnClick = Correo1Click
      end
      object Diasderecepcin1: TMenuItem
        Caption = 'Configurar dias de recepci'#243'n'
        OnClick = Diasderecepcin1Click
      end
    end
  end
end
