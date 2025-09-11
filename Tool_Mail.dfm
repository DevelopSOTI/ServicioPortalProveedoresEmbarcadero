object TM: TTM
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Configurar correo de la p'#225'gina'
  ClientHeight = 265
  ClientWidth = 337
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 321
    Height = 249
    Caption = 'Datos del correo del portal de proveedores'
    TabOrder = 0
    object Label1: TLabel
      Left = 65
      Top = 35
      Width = 37
      Height = 13
      Caption = 'Nombre'
    end
    object Label2: TLabel
      Left = 68
      Top = 62
      Width = 34
      Height = 13
      Caption = 'Asunto'
    end
    object Label3: TLabel
      Left = 33
      Top = 89
      Width = 69
      Height = 13
      Caption = 'Servidor SMTP'
    end
    object Label4: TLabel
      Left = 25
      Top = 116
      Width = 77
      Height = 13
      Caption = 'Puerto de salida'
    end
    object Label5: TLabel
      Left = 74
      Top = 143
      Width = 28
      Height = 13
      Caption = 'E-Mail'
    end
    object Label6: TLabel
      Left = 46
      Top = 170
      Width = 56
      Height = 13
      Caption = 'Contrase'#241'a'
    end
    object EDIT_NOMBRE: TEdit
      Left = 108
      Top = 32
      Width = 189
      Height = 21
      TabOrder = 0
    end
    object EDIT_ASUNTO: TEdit
      Left = 108
      Top = 59
      Width = 189
      Height = 21
      TabOrder = 1
    end
    object EDIT_SERVIDOR: TEdit
      Left = 108
      Top = 86
      Width = 189
      Height = 21
      TabOrder = 2
    end
    object EDIT_PUERTO: TEdit
      Left = 108
      Top = 113
      Width = 189
      Height = 21
      TabOrder = 3
    end
    object EDIT_CORREO: TEdit
      Left = 108
      Top = 140
      Width = 189
      Height = 21
      TabOrder = 4
    end
    object EDIT_PASSWORD: TEdit
      Left = 108
      Top = 167
      Width = 189
      Height = 21
      PasswordChar = '*'
      TabOrder = 5
    end
    object Button1: TButton
      Left = 108
      Top = 194
      Width = 189
      Height = 25
      Caption = 'Guardar cambios'
      TabOrder = 6
      OnClick = Button1Click
    end
  end
  object Conexion_MySQL: TADOConnection
    ConnectionString = 
      'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=soticomm_admon;PORT=330' +
      '6;DATABASE=soticomm_RelPro;SERVER=localhost;'
    LoginPrompt = False
    Mode = cmRead
    Left = 56
    Top = 208
  end
  object Select_Correo: TADOQuery
    Connection = Conexion_MySQL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM MAIL')
    Left = 160
    Top = 208
  end
  object Command_Update: TADOCommand
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 256
    Top = 208
  end
end
