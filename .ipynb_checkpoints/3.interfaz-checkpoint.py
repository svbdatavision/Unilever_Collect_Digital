import tkinter as tk
from tkinter import messagebox, ttk
from datetime import datetime
from PIL import Image, ImageTk  # pip install pillow

# Importamos clientes y template
import clientes.farmatodo as farmatodo
import clientes.olimpica as olimpica
from clientes.formato_template import exportar_template


def procesar_cliente(cliente):
    if cliente == "Farmatodo":
        return farmatodo.procesar()
    elif cliente == "Olimpica":
        return olimpica.procesar()
    else:
        raise ValueError("Cliente no soportado")


def ejecutar():
    cliente = combo_cliente.get()
    if not cliente:
        messagebox.showwarning("Atención ⚠️", "Seleccione un cliente antes de continuar.")
        return

    try:
        # Procesar datos según cliente
        df = procesar_cliente(cliente)

        # Fecha para el nombre del archivo de salida
        fecha = datetime.today().strftime("%Y%m%d")
        archivo_salida = f"Archivos/Template/Template_HRC_{cliente}_{fecha}.xlsx"

        # Exportar con formato estándar
        exportar_template(df, cliente, ruta_salida=archivo_salida)

        messagebox.showinfo(
            "Proceso finalizado ✅",
            f"El template para {cliente} fue generado correctamente.\n\n"
            f"Archivo exportado:\n{archivo_salida}"
        )
    except Exception as e:
        messagebox.showerror("Error ❌", f"Ocurrió un problema:\n{e}")


# --- Ventana principal ---
root = tk.Tk()
root.title("Generador de Template HRC")
root.geometry("500x400")

# --- Logo Unilever ---
try:
    logo = Image.open("Archivos/logo_unilever.png")
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
clientes = ["Farmatodo", "Olimpica"]
combo_cliente = ttk.Combobox(root, values=clientes, font=("Arial", 12), state="readonly")
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
