Para ver si esta instalado el servicio, buscarlo con los siguientes datos en los servicios de Windows
     	     Nombre = Servicio de sincronización
	Descripción = Servicio que actualiza la información del portal de proveedores


Metodo de inserción [registry]
	CLOSE_AUTO = False/True
	MAILS_SEND = False/True

Metodo de inserción [parametros]
	APLICA_DIR = FALSE/TRUE


Para ver los eventos del servicio ir al "Visor de eventos" de Windows,
en la carpeta "Registros de Windows" en "Aplicación", los eventos del servicio en la columna "Origen"
dice SyncService