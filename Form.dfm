object F: TF
  Left = 0
  Top = 0
  Caption = 'Servicio de sincronizaci'#243'n'
  ClientHeight = 329
  ClientWidth = 449
  Color = clBtnFace
  Constraints.MinHeight = 359
  Constraints.MinWidth = 465
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  DesignSize = (
    449
    329)
  PixelsPerInch = 96
  TextHeight = 13
  object Memo: TMemo
    Left = 8
    Top = 87
    Width = 433
    Height = 234
    Anchors = [akLeft, akTop, akRight, akBottom]
    Lines.Strings = (
      'Memo')
    ScrollBars = ssBoth
    TabOrder = 2
  end
  object ProgressBar: TProgressBar
    Left = 8
    Top = 55
    Width = 433
    Height = 26
    Anchors = [akLeft, akTop, akRight]
    TabOrder = 1
  end
  object Panel: TPanel
    Left = 8
    Top = 8
    Width = 433
    Height = 41
    Anchors = [akLeft, akTop, akRight]
    BevelOuter = bvNone
    Color = 3355443
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWhite
    Font.Height = -11
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
  end
end
