"""
Interfaz principal para generar los templates HRC
Funciona en Windows y Mac. Detecta automáticamente la carpeta de Excel relativa al ejecutable.
"""

import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
from PIL import Image, ImageTk

# --- Importar clientes ---
import clientes.farmatodo as farmatodo
import clientes.olimpica as olimpica

# --- Función para obtener la ruta base ---
def get_base_path():
    """
    Devuelve la carpeta donde se encuentra el ejecutable o el script.
    Permite que funcione tanto en PyInstaller como en modo script.
    """
    if getattr(sys, "frozen", False):
        macos_dir   = os.path.dirname(sys.executable)        # .../MyApp.app/Contents/MacOS
        contents_dir= os.path.dirname(macos_dir)            # .../MyApp.app/Contents
        app_bundle  = os.path.dirname(contents_dir)         # .../MyApp.app
        app_parent  = os.path.dirname(app_bundle)           # carpeta que contiene el .app (ej: /.../2.Automatizacion...)
        return app_parent
    return os.path.dirname(os.path.abspath(__file__))

# --- Función para procesar el cliente ---
def procesar_cliente(cliente):
    """
    Llama a la función procesar() del cliente seleccionado.
    """
    if cliente == "Farmatodo":
        return farmatodo.procesar()
    elif cliente == "Olimpica":
        return olimpica.procesar()
    else:
        raise ValueError("Cliente no soportado")

# --- Función del botón ---
def ejecutar():
    """
    Ejecuta el proceso de generación del template y muestra mensajes de éxito/error.
    """
    cliente = combo_cliente.get()
    if not cliente:
        messagebox.showwarning("Atención ⚠️", "Seleccione un cliente antes de continuar.")
        return

    try:
        base_path = get_base_path()
        os.chdir(base_path)  # Cambiamos el directorio para que los clientes lean los Excel

        resultado = procesar_cliente(cliente)

        # Fecha para el nombre del archivo
        fecha = datetime.today().strftime("%Y%m%d")
        archivo_salida = os.path.join(base_path, f"Template_HRC_{cliente}_{fecha}.xlsx")

        messagebox.showinfo(
            "Proceso finalizado ✅",
            f"El template para {cliente} fue generado correctamente.\n\n"
            f"Archivo exportado: {archivo_salida}"
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error ❌", f"Ocurrió un problema:\n{e}")

# --- Ventana principal ---
root = tk.Tk()
root.title("Generador de Template HRC")
root.geometry("500x400")

# --- Logo Unilever ---
try:
    base_path = get_base_path()
    logo_path = os.path.join(base_path, "Archivos/logo_unilever.png")
    logo = Image.open(logo_path)
    logo = logo.resize((120, 120))
    logo_tk = ImageTk.PhotoImage(logo)
    lbl_logo = tk.Label(root, image=logo_tk)
    lbl_logo.pack(pady=10)
except Exception as e:
    print("No se pudo cargar el logo:", e)

# Texto descriptivo
label = tk.Label(
    root,
    text="Seleccione el cliente para generar el Template HRC",
    font=("Arial", 12),
    pady=10
)
label.pack()

# Dropdown de clientes
clientes_list = ["Farmatodo", "Olimpica"]
combo_cliente = ttk.Combobox(root, values=clientes_list, font=("Arial", 12), state="readonly")
combo_cliente.set("")  # sin valor por defecto
combo_cliente.pack(pady=15)

# Botón de ejecutar
btn_generar = tk.Button(
    root, text="Generar Template", font=("Arial", 14), width=25, bg="#004C97", fg="white",
    command=ejecutar
)
btn_generar.pack(pady=20)

# Footer
footer = tk.Label(root, text="Unilever - Automatización HRC", font=("Arial", 9), fg="gray")
footer.pack(side="bottom", pady=5)

# Ejecutar ventana
root.mainloop()
