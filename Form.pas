unit Form;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, SvCom_Timer, Vcl.ComCtrls,
  Vcl.ExtCtrls;

type
  TF = class(TForm)
    Memo: TMemo;
    ProgressBar: TProgressBar;
    Panel: TPanel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  F: TF;

implementation

{$R *.dfm}

end.
