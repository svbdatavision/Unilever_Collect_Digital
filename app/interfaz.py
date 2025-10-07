"""
Interfaz principal para ejecutar el procesamiento de clientes.
Compatible con Windows y Mac. Detecta automáticamente la carpeta de ejecución.
"""

import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
from PIL import Image, ImageTk

# --- Importar módulos de clientes ---
from clientes.Colombia import olimpica, farmatodo, D1, copi, Cencosud, Euro

# --- Registro de clientes ---
CLIENTES = {
    "Olimpica": olimpica.procesar,
    "Farmatodo": farmatodo.procesar,
    "D1": D1.procesar,
    "Copidrogas": copi.procesar,
    "Cencosud": Cencosud.procesar,
    "Euro": Euro.procesar
}

# --- Obtener ruta base (para compatibilidad con PyInstaller) ---
def get_base_path():
    """
    Devuelve la carpeta donde se encuentra el ejecutable o el script.
    Funciona tanto en PyInstaller como en modo desarrollo.
    """
    if getattr(sys, "frozen", False):
        # Caso ejecutable
        base_path = os.path.dirname(sys.executable)
        # En Mac .app, hay que subir niveles hasta el bundle raíz
        if base_path.endswith("MacOS"):
            base_path = os.path.dirname(os.path.dirname(os.path.dirname(base_path)))
        return base_path
    else:
        # Caso script (.py)
        return os.path.dirname(os.path.abspath(__file__))

# --- Función para procesar cliente ---
def procesar_cliente(cliente):
    """
    Llama a la función procesar() correspondiente al cliente seleccionado.
    """
    if cliente not in CLIENTES:
        raise ValueError(f"Cliente '{cliente}' no está registrado.")
    return CLIENTES[cliente]()

# --- Acción al presionar el botón ---
def ejecutar():
    cliente = combo_cliente.get()
    if not cliente:
        messagebox.showwarning("Atención ⚠️", "Seleccione un cliente antes de continuar.")
        return

    try:
        base_path = get_base_path()
        os.chdir(base_path)

        df = procesar_cliente(cliente)

        # Guardar salida (opcional, si los scripts devuelven un DataFrame)
        fecha = datetime.today().strftime("%Y%m%d")
        archivo_salida = os.path.join(base_path, f"Resultado_{cliente}_{fecha}.xlsx")

        messagebox.showinfo(
            "Proceso completado ✅",
            f"El proceso de {cliente} finalizó correctamente.\n\n"
            f"Archivo exportado (si aplica):\n{archivo_salida}"
        )
        print(df.head())

    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error ❌", f"Ocurrió un problema:\n{e}")

# --- Crear ventana principal ---
root = tk.Tk()
root.title("Procesador de Clientes HRC")
root.geometry("480x400")
root.configure(bg="white")

# --- Cargar logo (si existe) ---
try:
    base_path = get_base_path()
    logo_path = os.path.join(base_path, "Archivos", "logo_unilever.png")
    if os.path.exists(logo_path):
        logo = Image.open(logo_path)
        logo = logo.resize((120, 120))
        logo_tk = ImageTk.PhotoImage(logo)
        lbl_logo = tk.Label(root, image=logo_tk, bg="white")
        lbl_logo.pack(pady=10)
except Exception as e:
    print("No se pudo cargar el logo:", e)

# --- Texto descriptivo ---
label = tk.Label(
    root,
    text="Seleccione el cliente que desea procesar:",
    font=("Arial", 12),
    bg="white"
)
label.pack(pady=10)

# --- Dropdown de clientes ---
combo_cliente = ttk.Combobox(
    root,
    values=list(CLIENTES.keys()),
    font=("Arial", 12),
    state="readonly"
)
combo_cliente.pack(pady=15)

# --- Botón principal ---
btn_procesar = tk.Button(
    root,
    text="Ejecutar Proceso",
    font=("Arial", 14),
    width=25,
    bg="#004C97",
    fg="white",
    command=ejecutar
)
btn_procesar.pack(pady=20)

# --- Footer ---
footer = tk.Label(
    root,
    text="Unilever - Automatización HRC",
    font=("Arial", 9),
    bg="white",
    fg="gray"
)
footer.pack(side="bottom", pady=10)

# --- Ejecutar ventana ---
root.mainloop()
