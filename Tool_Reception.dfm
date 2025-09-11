object TR: TTR
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Configurar dias de recepci'#243'n'
  ClientHeight = 241
  ClientWidth = 289
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
  object DBGridEmpresas: TDBGrid
    Left = 8
    Top = 8
    Width = 273
    Height = 225
    DataSource = DataSource_Empresas
    Options = [dgTitles, dgIndicator, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit, dgMultiSelect]
    PopupMenu = PopupMenu1
    ReadOnly = True
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridEmpresasDrawColumnCell
    Columns = <
      item
        Expanded = False
        FieldName = 'DIA_NUMERO'
        Visible = False
      end
      item
        Expanded = False
        FieldName = 'DIA_NOMBRE'
        Title.Caption = 'Nombre'
        Width = 125
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DIA_RECIBE'
        Title.Caption = 'Recepci'#243'n'
        Width = 82
        Visible = True
      end>
  end
  object Conexion_MySQL: TADOConnection
    ConnectionString = 
      'DRIVER=MySQL ODBC 5.3 Unicode Driver;UID=soticomm_admon;PORT=330' +
      '6;DATABASE=soticomm_RelPro;SERVER=localhost;'
    LoginPrompt = False
    Mode = cmRead
    Left = 40
    Top = 56
  end
  object Select_Dias: TADOQuery
    Connection = Conexion_MySQL
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      'SELECT * FROM DIAS'
      'ORDER BY DIA_NUMERO')
    Left = 120
    Top = 56
  end
  object DataSource_Empresas: TDataSource
    DataSet = Select_Dias
    Left = 208
    Top = 56
  end
  object Command_Update: TADOCommand
    Connection = Conexion_MySQL
    Parameters = <>
    Left = 40
    Top = 120
  end
  object PopupMenu1: TPopupMenu
    BiDiMode = bdLeftToRight
    MenuAnimation = [maTopToBottom]
    ParentBiDiMode = False
    Left = 120
    Top = 120
    object Autorizar1: TMenuItem
      Caption = 'Autorizar'
      OnClick = Autorizar1Click
    end
    object Bloquear1: TMenuItem
      Caption = 'Bloquear'
      OnClick = Bloquear1Click
    end
  end
end
