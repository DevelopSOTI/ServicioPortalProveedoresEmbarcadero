object M: TM
  OldCreateOrder = False
  OnCreate = NtServiceCreate
  DisplayName = 'Servicio de sincronizaci'#243'n'
  ShareProcess = True
  UseSynchronizer = False
  OnStart = NtServiceStart
  ServiceName = 'SyncService'
  Description = 'Servicio que actualiza la informaci'#243'n del portal de proveedores'
  FailureActions = <>
  Height = 237
  Width = 157
  StartedByScm = '0CBA3B41-40E4F531'
  object svWtsSessions: TsvWtsSessions
    SessionStateChanged = svWtsSessionsSessionStateChanged
    Left = 64
    Top = 96
  end
  object svLaunchFrontEnd: TsvLaunchFrontEnd
    CreateProcessFlags = []
    OnProcessLaunch = svLaunchFrontEndProcessLaunch
    OnProcessTerminate = svLaunchFrontEndProcessTerminate
    Left = 64
    Top = 48
  end
  object svTimer: TsvTimer
    OnTimer = svTimerTimer
    Left = 64
    Top = 144
  end
end
