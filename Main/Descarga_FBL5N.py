import win32com.client
import os
import time
import datetime


user = os.getenv("USERNAME")

def find_process(name):
    cObj = win32com.client.GetObject("winmgmts://.")
    processes = cObj.ExecQuery(f"SELECT * FROM Win32_Process WHERE Name = '{name}'")
    return len(processes) > 0

def esperar(segundos):
    time.sleep(segundos)


def get_p1p_session():
     try:
         sap_gui = win32com.client.GetObject("SAPGUI")
         appl = sap_gui.GetScriptingEngine
         for i in range(appl.Children.Count):
             connection = appl.Children(i)
             if "P1P" in connection.Description:
                 print("Conexión P1P ya activa.")
                 return connection.Children(0)
     except Exception as e:
         print(f"No se pudo obtener SAPGUI: {e}")
     return None

def formato_mes(fecha=None):
    if fecha is None:
        fecha = datetime.datetime.now()
    meses = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
        5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
        9: "Setiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }

    mes_nombre = meses[fecha.month]
    año = fecha.year
    return f"{mes_nombre}_{año}"
 
def cerrar_archivo_excel(nombre_archivo):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        libros_a_cerrar = []

        # Buscar libros que coincidan
        for workbook in list(excel.Workbooks):
            if nombre_archivo.lower() in workbook.Name.lower():
                libros_a_cerrar.append(workbook)

        # Cerrar los libros encontrados
        for libro in libros_a_cerrar:
            nombre = libro.Name  # Guardar el nombre antes de cerrar
            libro.Close(SaveChanges=False)
            print(f"El archivo de Excel '{nombre}' ha sido cerrado.")

        # Cerrar Excel si no quedan libros abiertos
        if excel.Workbooks.Count == 0:
            excel.Quit()

    except Exception as e:
        print(f"Error al cerrar Excel: {e}")

def descargar_datos():
    # Calcular fechas relevantes
    fecha_hoy = f"{time.strftime('%d')}.{time.strftime('%m')}.{time.strftime('%Y')}"

# Verificar si SAP Logon está abierto
    if not find_process("saplogon.exe"):
        print("SAP Logon no está abierto. Iniciando...")
        os.system("start /min saplogon.exe")
        esperar(5)
    else:
        print("SAP Logon ya está abierto.")

# Verificar si ya hay una sesión P1P activa
    session = get_p1p_session()

# Si no hay sesión P1P, abrirla
    if session is None:
        print("Abriendo nueva conexión a P1P...")
        sap_gui = win32com.client.GetObject("SAPGUI")
        appl = sap_gui.GetScriptingEngine
        connection = appl.OpenConnection("P1P - Cordillera - ECC Production", True)
        session = connection.Children(0)

    # Confirmación
    print("Sesión SAP lista para usar.")

    # Descarga de transacciones
    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "/NFBL5N"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/chkX_SHBV").selected = True
    session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = ""
    session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "9370"
    session.findById("wnd[0]/usr/ctxtPA_STIDA").text = fecha_hoy
    session.findById("wnd[0]/usr/ctxtPA_VARI").text = "/LADR"
    session.findById("wnd[0]/usr/ctxtPA_VARI").setFocus()
    session.findById("wnd[0]/usr/ctxtPA_VARI").caretPosition = 5
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
    session.findById("wnd[1]").sendVKey(0)
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = fr"C:\Users\{user}\OneDrive - Unilever\Documents\SAP\SAP GUI\A"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = fr"FBL5N_{fecha_hoy}.XLSX"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
    session.findById("wnd[1]/tbar[0]/btn[11]").press()


# Ejecutar la función principal
descargar_datos()


    
    


