Para ver si esta instalado el servicio, buscarlo con los siguientes datos en los servicios de Windows
     	     Nombre = Servicio de sincronizaci�n
	Descripci�n = Servicio que actualiza la informaci�n del portal de proveedores


Metodo de inserci�n [registry]
	CLOSE_AUTO = False/True
	MAILS_SEND = False/True

Metodo de inserci�n [parametros]
	APLICA_DIR = FALSE/TRUE


Para ver los eventos del servicio ir al "Visor de eventos" de Windows,
en la carpeta "Registros de Windows" en "Aplicaci�n", los eventos del servicio en la columna "Origen"
dice SyncService