unit Main;

interface

uses
    Windows, Messages, SysUtils, Classes, Graphics, Dialogs, Controls, ExtCtrls, ActiveX, StrUtils, DateUtils,
    SvCom_NTService, SvCom_Timer, SvCom_LaunchFrontEnd, SvCom_WtsSessions, SvCom_WtsApi;

type
  TM = class(TNtService)
    svWtsSessions: TsvWtsSessions;
    svLaunchFrontEnd: TsvLaunchFrontEnd;
    svTimer: TsvTimer;
    procedure NtServiceCreate(Sender: TObject);
    procedure NtServiceStart(Sender: TNtService; var DoAction: Boolean);
    procedure svLaunchFrontEndProcessLaunch(Sender: TObject; Process: TsvLaunchedProcess);
    procedure svLaunchFrontEndProcessTerminate(Sender: TObject; Process: TsvLaunchedProcess);
    procedure svTimerTimer(Sender: TObject);
    procedure svWtsSessionsSessionStateChanged(Sender: TObject; Session: TsvWtsSessionInfo);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  M: TM;

implementation

uses
  Data, Form, Func;

{$R *.DFM}

{$REGION 'NtServiceCreate - EVENTO CREATE, LE ESPECIFICO CUAL APLICACIÓN QUIERO QUE EJECUTE EL SERVICIO'}
procedure TM.NtServiceCreate(Sender: TObject);
begin
  svLaunchFrontEnd.ApplicationName := ExtractFilePath(ParamStr(0)) + 'Viewer.exe';
  svLaunchFrontEnd.CurrentDirectory := ExtractFilePath(ParamStr(0));

  // SI MODE_APPLI ES 'F' QUIERE DECIR QUE SE EJECUTARA EN MODO VISUAL
  if (D.MODE_APPLI = 'F') then
    begin
      D.svTimerSync.Enabled := True;
    end;
end;
{$ENDREGION}

{$REGION 'NtServiceStart - CUANDO INICIA SESIÓN INICIA LA APLICACIÓN DE ESCRITORIO'}
procedure TM.NtServiceStart(Sender: TNtService; var DoAction: Boolean);
  var
    Session: TsvWtsSessionInfo;
    i: Integer;
begin
  svWtsSessions.UpdateSessions;
  svWtsSessions.Lock;

  try
    for i := 0 to svWtsSessions.SessionCount - 1 do
      begin
        Session := svWtsSessions.Sessions[i];

        if (Session.State <> WTSActive) then
          begin
            Continue;
          end;

        svLaunchFrontEnd.Launch(Session.SessionId);
      end;
  finally
    svWtsSessions.Unlock;
  end; // }

  // SI MODE_APPLI ES 'S' QUIERE DECIR QUE SE EJECUTARA EN MODO DE SERVICIO
  if (D.MODE_APPLI = 'S') then
    begin
      D.svTimerSync.Enabled := True;
    end;
end;
{$ENDREGION}

{$REGION 'svLaunchFrontEndProcessLaunch - MANDA UN EVENTO CUANDO INICIA LA APLICACIÓN'}
procedure TM.svLaunchFrontEndProcessLaunch(Sender: TObject; Process: TsvLaunchedProcess);
begin
  // // EventLog.LogMessage('Sincronizador ejecutado el dia ' + FormatDateTime('dd/mm/yyyy', Now) + ' a las ' + FormatDateTime('hh:nn:ss', Now));
end;
{$ENDREGION}

{$REGION 'svLaunchFrontEndProcessTerminate - SI LA APLICACIÓN SE CIERRA, EL SERVICIO LA VUELVE A INICIAR'}
procedure TM.svLaunchFrontEndProcessTerminate(Sender: TObject; Process: TsvLaunchedProcess);
  var
    Session: TsvWtsSessionInfo;
begin
  svWtsSessions.UpdateSessions;
  svWtsSessions.Lock;

  try
    Session := svWtsSessions.SessionsById[Process.SessionId];

    if not Assigned(Session) or (Session.State <> WTSActive) then
      begin
        Process.Free;
        Exit;
      end;
  finally
    svWtsSessions.Unlock;
  end;

  svLaunchFrontEnd.Launch(Process.SessionId); // }
end;
{$ENDREGION}

{$REGION 'svTimerTimer - TIMER PARA ACTUALIZAR LAS SESIONES'}
procedure TM.svTimerTimer(Sender: TObject);
begin
  svWtsSessions.UpdateSessions;
end;
{$ENDREGION}

{$REGION 'svWtsSessionsSessionStateChanged - DETECTA CUANDO LA SESION CAMBIO DE ESTATUS'}
procedure TM.svWtsSessionsSessionStateChanged(Sender: TObject; Session: TsvWtsSessionInfo);
  var
    LP: TsvLaunchedProcess;
    i: Integer;
begin
  if (Session.State <> WTSActive) then
    begin
      Exit;
    end;

  for i := svLaunchFrontEnd.LaunchedProcessCount - 1 downto 0 do
    begin
      LP := svLaunchFrontEnd.LaunchedProcess[i];

      if not LP.IsAlive then
        begin
          LP.Free;
          Continue;
        end;

      if LP.SessionId = Session.SessionId then
        begin
          Exit;
        end;
    end;

  svLaunchFrontEnd.Launch(Session.SessionId); // }
end;
{$ENDREGION}

end.



