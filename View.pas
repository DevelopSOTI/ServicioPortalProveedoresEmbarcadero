unit View;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ComCtrls,
  Vcl.ExtCtrls, Vcl.AppEvnts, Vcl.Menus, WinSvc, ShellApi, Jpeg;

type
  TV = class(TForm)
    PopupMenu: TPopupMenu;
    Abrirsincronizador1: TMenuItem;
    Configurador1: TMenuItem;
    Salir1: TMenuItem;
    TrayIcon: TTrayIcon;
    ApplicationEvents: TApplicationEvents;
    Timer: TTimer;
    Tittle: TPanel;
    LastAuto: TPanel;
    ProgressBar: TProgressBar;
    MemoErrors: TMemo;
    Image: TImage;
    procedure FormCreate(Sender: TObject);
    procedure FormActivate(Sender: TObject);
    procedure ApplicationEventsMinimize(Sender: TObject);
    procedure TrayIconDblClick(Sender: TObject);
    procedure TimerTimer(Sender: TObject);
    procedure Abrirsincronizador1Click(Sender: TObject);
    procedure Configurador1Click(Sender: TObject);
    procedure Salir1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  V: TV;

implementation

{$R *.dfm}

{$REGION 'IS_RUNNING - REVISA SI YA HAY UNA APLICACIÓN CORRIENDO'}
Function IS_RUNNING( aplicacion :string ):Boolean;
  var
    MiMutex :THandle;
begin
  MiMutex := CreateMutex( nil, True, PChar( aplicacion ) );
  if ( MiMutex = 0 ) then
    begin
      Result := False;
    end
  else
    begin
      if ( GetLastError = ERROR_ALREADY_EXISTS ) then
        begin
          Result := True;
        end
      else
        begin
          Result := False;
        end;
    end;
end;
{$ENDREGION}

{$REGION 'EJECUTAR - EJECUTA UNA APLICACIÓN EXTERNA'}
Function EJECUTAR( Programa :string; Esperar :Boolean = True ):Boolean;
  var
    ProcInfo :TProcessInformation;
    Info :TStartupInfo;
begin
  FillChar( Info, SizeOf( Info ), 0 );
  Info.cb := SizeOf( Info );
  Info.dwFlags := STARTF_USESHOWWINDOW;
  Info.wShowWindow := SW_SHOWNORMAL;
  Result := CreateProcess( PChar( Programa ), PChar( Programa ), nil, nil, False, 0, nil, nil, Info, ProcInfo );
  if Esperar then
    begin
      while WaitForSingleObject( ProcInfo.hProcess, 100 ) = WAIT_TIMEOUT do
        begin
          Application.ProcessMessages;
          if Application.Terminated then
            begin
              Break;
            end;
        end;
    end;
end;
{$ENDREGION}



{$REGION 'FormCreate - EVENTO CREATE'}
procedure TV.FormCreate(Sender: TObject);
  var
    MyIcon :TIcon;
begin
  ProgressBar.MarqueeInterval := 1;
  ProgressBar.Step := 1;

  TrayIcon.Icons := TImageList.Create( Self );
  MyIcon := TIcon.Create;

  TrayIcon.Icon.Assign( MyIcon );

  { MyIcon.LoadFromFile( 'Tray Icon/reload.ico' );
  TrayIcon.Icons.AddIcon( MyIcon ); // }

  { MyIcon.LoadFromFile( 'Tray Icon/Tray_E.ico' );
  TrayIcon.Icons.AddIcon( MyIcon ); // }

  { MyIcon.LoadFromFile( 'Tray Icon/Tray_D.ico' );
  TrayIcon.Icons.AddIcon( MyIcon ); // }

  TrayIcon.Hint := 'Sincronizador para el portal de proveedores';
  TrayIcon.AnimateInterval := 200;

  TrayIcon.BalloonTitle := 'Sincronizador para el portal de proveedores';
  TrayIcon.BalloonFlags := bfInfo;

  TrayIcon.Visible := True;
end;
{$ENDREGION}

{$REGION 'FormActivate - EVENTO ACTIVATE'}
procedure TV.FormActivate(Sender: TObject);
begin
  if IS_RUNNING( 'Viewer.exe' ) then
    begin
      Close
    end;

  // Application.Minimize;
  Timer.Enabled := True;
end;
{$ENDREGION}

{$REGION 'ApplicationEventsMinimize - EVENTO DE MINIMIZAR'}
procedure TV.ApplicationEventsMinimize(Sender: TObject);
begin
  Hide();
  WindowState := wsMinimized;

  // // TrayIcon.Visible := True;
  // TrayIcon.Animate := True;

  // TrayIcon.BalloonHint := 'Doble clic para restaurar';
  // TrayIcon.ShowBalloonHint;
end;
{$ENDREGION}

{$REGION 'EVENTOS REFERENTES AL TRAY ICON'}
procedure TV.TrayIconDblClick(Sender: TObject);
begin
  Show();
  WindowState := wsNormal;
  Constraints.MinHeight := 351;
  Constraints.MinWidth := 505;
  Application.BringToFront();
end;

procedure TV.Abrirsincronizador1Click(Sender: TObject);
begin
  Show();
  WindowState := wsNormal;
  Application.BringToFront();
end;

procedure TV.Configurador1Click(Sender: TObject);
begin
  EJECUTAR( 'Tools.exe', False );
end;

procedure TV.Salir1Click(Sender: TObject);
begin
  V.Close;
end;
{$ENDREGION}

{$REGION 'TimerTimer - CADA DETERMINADO TIEMPO HACE LA ACTUALIZACIÓN'}
procedure TV.TimerTimer(Sender: TObject);
  var
    ProgressMax, Row :Integer;
    PARAMETER, RutaFichero :string;
    F :TextFile;
begin
  RutaFichero := ExtractFilePath( ParamStr( 0 ) ) + 'EventLog/EventLog.ini';
  Timer.Enabled := False;

  try
    if ( FileExists( RutaFichero ) ) then
      begin
        PARAMETER := '';
        Row := 1;

        AssignFile( F, RutaFichero );
        Reset( F );
        while not Eof( F ) do
          begin
            Readln( F, PARAMETER );

            if ( PARAMETER <> '' ) then
              begin
                case Row of
                  1:begin
                      ProgressMax := StrToInt( PARAMETER );
                      ProgressBar.Max := ProgressMax;
                    end;
                  2: ProgressBar.Position := StrToInt( PARAMETER );
                  3: Tittle.Caption := PARAMETER;
                  4: LastAuto.Caption := PARAMETER;
                  5:begin
                      MemoErrors.Lines.Add( PARAMETER );
                      MemoErrors.Lines.Add( '' );
                    end;
                  6:begin
                      if ( PARAMETER = 'I' ) then // INICIA
                        begin
                          TrayIcon.Animate := True;
                          MemoErrors.Lines.Clear;
                        end;
                      if ( PARAMETER = 'F' ) then // FINALIZA
                        begin
                          TrayIcon.Animate := False;
                          if ( MemoErrors.Text <> '' ) then // SI HUBO ERRORES LOS GUARDA EN SERVICE REPORT
                            begin
                              MemoErrors.Lines.SaveToFile( 'Service Report/Service Report ' + FormatDateTime( 'yyyy-mm-dd hhnnss', Now ) + '.txt' );
                            end;
                        end;
                    end;
                end;
              end;

            Inc( Row );
          end;
        CloseFile( F );
        DeleteFile( RutaFichero );
      end;
  except
  end;

  V.Update;
  Timer.Enabled := True;
end;
{$ENDREGION}

end.
