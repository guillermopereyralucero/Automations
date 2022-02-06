#########################################################################################################
#PAQUETES/LIBRERIAS
#########################################################################################################
import pathlib  #Por defecto toma de donde esté instalado Python, por ej. C:/Python/...
import psutil #Para cerrar App Outlook
import sys  #Para cerrar luego de correr
import schedule   #Para programar rutina
import win32com.client as client #Para aplicaciones, en este caso Outlook
from SQLTodos import SQLTodos #Class que permite conexión con SQL
from os import * #Para obtener procesos Windows

from openpyxl import Workbook				 #Para trabajar libro
from openpyxl import load_workbook			#Para cargar libro
from openpyxl import cell					 #Para operar con celdas
import xlrd					  #Para operar con archivos .xls (viejos)

from datetime import * #Para gestión de fechas
import logging  #Para log, debug y errores
import shutil #Para mover/copiar archivos/carpetas

import pyautogui as PAG   #Para automatizar
import time as TM #Gestión de tiempos, esperas y demoras
import tkinter as tk    #Para portapapeles
import pyperclip as PC  #Para portapapeles

#########################################################################################################
#LOG
#########################################################################################################
print(datetime.now().strftime('%m-%d %H:%M:%S:%f'),"- Inicio de programa.")
log=f'F:/ProgramasPython/CorreoSdaLinea/Logs/CSL_{datetime.now().strftime("%Y-%m-%d")}.txt'
print(f'Loggin to file: {log}')
check=makedirs(path.dirname(log),exist_ok=True)
logger=logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)
formatter=logging.Formatter('\n%(asctime)s - %(levelname)s - at line: %(lineno)d - %(message)s')
file_handler=logging.FileHandler(log)
stream_handler = logging.StreamHandler()
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)
stream_handler.setFormatter(formatter)
logger.addHandler(file_handler)
logger.addHandler(stream_handler)
def plog(notas):
	print(datetime.now().strftime('%m-%d %H:%M:%S:%f'),"- "+str(notas))
	logger.debug(str(notas))
plog("Log de errores creado correctamente.")

#########################################################################################################
#DEFINICIONES PREVIAS
#########################################################################################################

def correo(destino,cc,bcc,asunto,cuerpoPlano,cuerpoHTML,adjunto):
#OUTLOOK - INICIO
	plog("Inicio de Outlook.") 
#OUTLOOK - CIERRA PROCESO OUTLOOK
	for process in psutil.process_iter():	   #Recorre todos los procesos
		if process.name() == "OUTLOOK.EXE":	 #Busca el proceso de la App por nombre
			system("TASKKILL /F /IM " + str(process.name()) + " /T")	   #Manda el TaskKill por CMD
	plog("Procesos anteriores de Outlook cerrados correctamente.")	 
#OUTLOOK - COMIENZA APP
	startfile("outlook") #Inicia App Outlook
	sleep(20)  #Espera intervalo (en segundos) para que cargue la App
	outlook = client.Dispatch("Outlook.Application")
	sleep(5)  #Espera intervalo (en segundos) para que cargue la App
	plog("Outlook abierto correctamente.")
	namespace=outlook.GetNamespace("MAPI")
	sleep(2)
	message=outlook.CreateItem(0) #Abre un nuevo mensaje
#OUTLOOK - CABECERA
	message.To=destino #Acá irá luego la lista de destinararios
	sleep(2)
	message.Display()   #Muestra en pantalla la redaccion del mensaje 
	if cc!=None:
		message.CC=cc #Destinatarios en Copia
	sleep(5)
	if bcc!=None:
		message.BCC=bcc #Destinatarios en Copia Oculta
	if asunto!=None:
		message.Subject=asunto #"Segunda Linea - Ingresos" #Asunto
	plog("Cabecera cargada correctamente.")
#OUTLOOK - CUERPO HTML
	if message.Body==None:
		message.HTMLBody=cuerpoHTML	#ingresa el cuerpo HTML 
	else: message.Body=cuerpoPlano  #ingresa el cuerpo plano
	sleep(10)  
	if adjunto!=None:
		message.Attachments.Add(adjunto) #Carga archivo adjunto, acá irá la ruta
	sleep(10)
	plog("Cuerpo y adjunto cargados correctamente.")
#OUTLOOK - ENVIO
	message.Save()	#Guarda mensaje
	sleep(3)  
	message.Send()	#Envia mensaje
	outbox=namespace.GetDefaultFolder(4)
	outbox=int(outbox.Items.Count)
	while outbox!=0:
	  namespace.SendAndReceive(True)
	  plog("Esperando 'EnviarYRecibir'...")
	  sleep(5)
	  outbox=namespace.GetDefaultFolder(4)
	  outbox=int(outbox.Items.Count)
	sleep(5)
	# outlook.quit()
	plog("El correo fue enviado correctamente.") #Una vez que la bandeja de salida está en cero, indica que el correo se envió correctamente.


def correoSdaLinea():
#########################################################################################################
#SQL - CHEQUEO PREVIO
#########################################################################################################
#SQL - CHEQUEO PREVIO - INICIO DE SESION SERVER
	SERVER = SQLTodos('SERVER','USER','PASS','DATABASE')
	global_app = None
	pid = getpid()
	plog("SQL SERVER conectado correctamente.")
#SQL - CHEQUEO PREVIO - SERVER
	verificacionSqlLista=list(SERVER.Query(f"SELECT CSL FROM [SERVER].[DATABASE].[dbo].[correoAutomaticoCalendarioControl] WHERE fecha=CAST(GETDATE() AS DATE)"))
	verificacionSql=verificacionSqlLista[0][0]
	if verificacionSql==0:
		plog("Según calendario SQL, hoy no se envía el correo.")
	elif verificacionSql==None:
		plog("Revisar Calendario SQL: número de proceso NULL o erroneo.")
	elif verificacionSql>1:
		plog("Según calendario SQL, el correo fue enviado anteriormente.")
	elif verificacionSql==1:
		intentoSql=SERVER.Query(f"UPDATE [SERVER].[DATABASE].[dbo].[correoAutomaticoCalendarioControl] SET [CSL_intento] += 1 WHERE fecha=CAST(GETDATE() AS DATE)")
		plog("Intentos SQL: "+str(intentoSql))
#########################################################################################################
#CHEQUEO ARCHIVO - VERIFICA EXISTENCIA 
#########################################################################################################
		ayer=(datetime.now() - timedelta(1)).strftime('%Y-%m-%d')
		archiOri=f'//SERVER2/d/ingresosSdaLinea/IngresosSdaLinea.xls'
		archiDes=f'//SERVER2/d/ingresosSdaLinea/envios/IngresosSdaLinea-{ayer}.xls'
		plog(str(path.exists(archiOri)))
		if path.exists(archiOri):
#########################################################################################################
#CHEQUEO ARCHIVO - VERIFICA FECHA DE CELDAS
#########################################################################################################
			#CARGA EXCEL EXISTENTE
			libro=xlrd.open_workbook(archiOri)										 
			hoja=libro.sheet_by_name("SdaLinea")
			#OBTIENE CONTENIDO CELDA
			celda=str(hoja.cell(1,0).value)
			fechaCelda=celda[0:10]
			plog(fechaCelda)
			#CHEQUEO FECHA CELDA
			if ayer==fechaCelda:
				shutil.copy(archiOri,archiDes)
				sleep(10)
#########################################################################################################
#OUTLOOK - ENVIAR MAIL DE INFORME
#########################################################################################################
				destino='' #Aquí van los destinatarios del informe
				cc='' #Aquí van los destinatarios del informe en copia, de no ir queda None
				bcc=None #Reemplazar el None por los destinatarios a quienes se quiera enviar en copia oculta.
				asunto="Ingresos - Segunda Linea" #Aquí reemplazar por el asunto.
				cuerpoPlano="Buenos días, en la hoja 'SdaLinea' del archivo adjunto, se encuentran los ingresos por usuarios de segunda linea y en la otra hoja estan los demas.\n\nAnte cualquier inconveniente, favor de consultar a las casillas que se encuentran en copia de este correo.\n\nMuchas gracias.\n\nSaludos." 
				cuerpoHTML=None #Reemplazar cuerpoPlano o cuerpoHTML según corresponda.
				adjunto=archiDes	#Definir archivo a adjuntar con variable archiDes
				correo(destino,cc,bcc,asunto,cuerpoPlano,cuerpoHTML,adjunto)	#Ejecutar la función correo con los parametros/variables establecidas previamente.
#########################################################################################################
#SQL - INICIO DE SESION SERVER
#########################################################################################################
				SERVER=SQLTodos('SERVER','USER','PASS','DATABASE')
				global_app=None
				pid=getpid()
				plog("SQL SERVER conectado correctamente.")
#########################################################################################################
#SQL - ACTUALIZA NUMERO PROCESO EN TABLA DE CONTROL SQL
#########################################################################################################
				controlSql=(SERVER.Query(f"UPDATE [SERVER].[DATABASE].[dbo].[correoAutomaticoCalendarioControl] SET CSL=[CSL_intento] WHERE fecha=CAST(GETDATE() AS DATE)"))
				plog(str(controlSql))
				plog("Se actualizó control de SQL SERVER")
#########################################################################################################
#SQL - INICIO DE SESION [SERVER2]
#########################################################################################################
				_SERVER2=SQLTodos('SERVER2','USER','PASS','DATABASE')
				global_app=None
				pid=getpid()
				plog("SQL SERVER2 conectado correctamente.")
#########################################################################################################
#SQL - ACTUALIZA TABLA "actualizacion" que controla varios procesos e informes.
#########################################################################################################
				controlSql=(_SERVER2.Query(f"UPDATE [DATABASE2].[dbo].[actualizacion] SET nota='Correo enviado: '+CONVERT(VARCHAR,GETDATE(),21) WHERE tabla='ingreso2daLinea'"))
				plog(str(controlSql))
				plog("Se actualizó tabla control de SQL [SERVER2].[DATABASE2].dbo.actualizacion")
#########################################################################################################        
#OUTLOOK - ARCHIVO DESACTUALIZADO				
#########################################################################################################
			else:
				plog("Archivo con datos desactualizados.")
				#ENVIAR MAIL DE ALARMA
				destino='' #Reemplazar por destinatarios a los cuales avisar en caso de que el proceso tenga algún problema, en este caso, el archivo a adjuntar se encuentre desactualizado.
				cc='' #Reemplazar por destinatarios en copia.
				bcc=None #Reemplazar por destinatarios en copia oculta.
				asunto="Ingresos - Segunda Linea - Archivo con datos desactualizados" #Reemplazar por asunto
				cuerpoPlano="Archivo \\\\SERVER2\\D\\ingresosSdaLinea\\IngresosSdaLinea.xls con datos desactualizados. Verificar:\n\nSV: SERVER2\nDB: [DATABASE2]\nJB: [DATABASE2]_ingresosSegundaLinea\nSP: [SERVER2].[DATABASE2].dbo.get_ingresosSdaLinea\nTB: [SERVER2].[DATABASE2].dbo.ingresoSdaLinea\n\nNota: el programa detectó que la fecha de la celda A:2 de la hoja 'SdaLinea' no coincide con la fecha de ayer. \n\nSaludos."
				cuerpoHTML=None #Reemplazar por observación cuerpoPlano o cuerpoHTML segun se prefiera.
				adjunto=None 
				correo(destino,cc,bcc,asunto,cuerpoPlano,cuerpoHTML,adjunto)
#########################################################################################################
#OUTLOOK - ARCHIVO INEXISTENTE				
#########################################################################################################
		else:
			plog("Archivo no encontrado.")
			#ENVIAR MAIL DE ALARMA
			destino='' #Reemplazar por destinatarios a los cuales avisar en caso de que el proceso tenga algún problema, en este caso, el archivo a adjuntar no exista en el directorio.
			cc='' #Reemplazar por destinatarios en copia.
			bcc=None #Reemplazar por destinatarios en copia oculta.
			asunto="Ingresos - Segunda Linea - Archivo no encontrado." 
			cuerpoPlano="Archivo \\\\SERVER2\\D\\ingresosSdaLinea\\IngresosSdaLinea.xls no encontrado. Verificar SQL y tabla de actualización:\n\nSV: SERVER2\nDB: [DATABASE2]\nJB: [DATABASE2]_ingresosSegundaLinea\nSP: [SERVER2].[DATABASE2].dbo.get_ingresosSdaLinea\nTB: [SERVER2].[DATABASE2].dbo.ingresoSdaLinea\n\nQuery: SELECT * FROM [SERVER2].[DATABASE2].[dbo].[actualizacion] WHERE tabla='ingreso2daLinea'\n\n Saludos."
			cuerpoHTML=None	#Reemplazar por observación cuerpoPlano o cuerpoHTML segun se prefiera.
			adjunto=None
			correo(destino,cc,bcc,asunto,cuerpoPlano,cuerpoHTML,adjunto)
#########################################################################################################
#OUTLOOK - CALENDARIO SQL CONTROL MAL CONFIGURADO 
#########################################################################################################
	else:
		plog("Revisar Calendario SQL: número de proceso NULL o erroneo.")
		#ENVIAR MAIL A NOSOTROS
		destino='' #Reemplazar por destinatarios a los cuales avisar en caso de que el calendario control esté mal configurado. 
		cc='' #Reemplazar por destinatarios en copia
		bcc=None #Reemplazar por destinatarios en copia oculta
		asunto="Ingresos - Segunda Linea - Revisar Calendario SQL."
		cuerpoPlano="Revisar Calendario SQL. Verificar:\n\nSV: SERVER\nDB: DATABASE\nTB: [SERVER].DATABASE.dbo.correoAutomaticoCalendarioControl\n\nQuery: SELECT fecha,CSL FROM [SERVER].[DATABASE].[dbo].[correoAutomaticoCalendarioControl] WHERE fecha=CAST(GETDATE() AS DATE)\nNota: Para que el correo se envíe determinado día, el campo CSL tiene que ser CSL=1.\n\n Saludos."
		cuerpoHTML=None	#Reemplazar por observación cuerpoPlano o cuerpoHTML segun se prefiera.
		adjunto=None
		correo(destino,cc,bcc,asunto,cuerpoPlano,cuerpoHTML,adjunto)

#########################################################################################################
#EJECUCIÓN DE JOB
#########################################################################################################
try:
	horaInicio=datetime.now()
	correoSdaLinea()
	horaFin=datetime.now()
	duracion=str(horaFin-horaInicio)
	plog('Duración: '+duracion)
except: 
	plog("Correo no enviado - Error en el programa, revisar log.")
	#ENVIAR MAIL DE ALARMA
	destino='' #Reemplazar por destinatarios a los cuales avisar en caso de que el programa falle.
	cc='' #Reemplazar por destinatarios en copia
	bcc=None #Reemplazar por destinatarios en copia oculta
	asunto="Ingresos - Segunda Linea - Correo no enviado - Error en el programa, revisar log"
	cuerpoPlano="Verificar último log en \\\\SERVIDOR3\\F:\\ProgramasPython\\CorreoSdaLinea\\Logs . Verificar:\n\nSV: SERVER2\nDB: [DATABASE2]\nJB: [DATABASE2]_ingresosSegundaLinea\nSP: [SERVER2].[DATABASE2].dbo.get_ingresosSdaLinea\nTB: [SERVER2].[DATABASE2].dbo.ingresoSdaLinea\n\nNota: Chequear PC SERVIDOR3 (user: usuario , pass: contraseña), Outlook y recursos. \n\nSaludos."
	cuerpoHTML=None
	adjunto=None
	correo(destino,cc,bcc,asunto,cuerpoPlano,cuerpoHTML,adjunto)
