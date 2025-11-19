import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import ImageTk, Image
import os

from Peru import Template_Cencosud, Template_mayorsa, Template_NF, Template_SPSA, Template_Tottus
from Colombia import Cencosud, copi, D1, farmatodo, olimpica, Euro, oxxo
from Ecuador import El_Rosado, Favorita, FARCOMED, DIFARE, Mega_Santa_Maria, FARMAENLACE

archivos_remittance = []
archivo_fbl5n = None
pais_actual = None

paises = ["Peru", "Colombia", "Ecuador"]

clientes_por_pais = {
    "Peru": ["Cencosud Peru", "Mayorsa", "Nortfarma", "SPSA", "Tottus"],
    "Colombia": ["Cencosud Colombia", "Copidrogas", "D1", "Farmatodo", "Olimpica", "Euro", "Oxxo"],
    "Ecuador": ["Favorita", "FARCOMED", "DIFARE", "Mega Santa Maria", "Farmaenlace", "El Rosado"]
}

procesadores = {
    "Cencosud Peru": Template_Cencosud,
    "Mayorsa": Template_mayorsa,
    "Nortfarma": Template_NF,
    "SPSA": Template_SPSA,
    "Tottus": Template_Tottus,
    "Cencosud Colombia": Cencosud,
    "Copidrogas": copi,
    "D1": D1,
    "Farmatodo": farmatodo,
    "Olimpica": olimpica,
    "Favorita": Favorita,
    "FARCOMED": FARCOMED,
    "Mega Santa Maria": Mega_Santa_Maria,
    "Farmaenlace": FARMAENLACE,
    "DIFARE": DIFARE,
    "El rosado": El_Rosado
}

def seleccionar_remittance():
    global archivos_remittance
    archivos_remittance = filedialog.askopenfilenames(title="Seleccionar archivos Remittance")
    remittance_status.config(text=f"{len(archivos_remittance)} archivo(s) Remittance cargado(s)")

def seleccionar_fbl5n():
    global archivo_fbl5n
    archivo_fbl5n = filedialog.askopenfilename(title="Seleccionar archivo FBL5N")
    fbl5n_status.config(text="Archivo FBL5N cargado")

def actualizar_clientes(event=None):
    global pais_actual
    pais_actual = pais_var.get()
    import config
    config.pais_actual = pais_actual  # ← Actualiza el módulo config
    clientes = clientes_por_pais.get(pais_actual, [])
    cliente_dropdown['values'] = clientes
    cliente_dropdown.set('')
    fbl5n_status.config(text="Haz clic para adjuntar el FBL5N")
    remittance_status.config(text="Haz clic para adjuntar los Remittance")

def actualizar_visibilidad_fbl5n(event=None):
    cliente = cliente_var.get()
    if cliente in []:
        fbl5n_button.config(state="disabled")
    else:
        fbl5n_button.config(state="normal")

def procesar_cliente():
    cliente = cliente_var.get()
    if not archivos_remittance:
        messagebox.showwarning("Advertencia", "Por favor, adjunta al menos un archivo Remittance.")
        return
    if cliente not in [] and not archivo_fbl5n:
        messagebox.showwarning("Advertencia", "Este cliente requiere también el archivo FBL5N.")
        return
    try:
        total_exportados = 0
        for archivo in archivos_remittance:
            resultado = procesadores[cliente].procesar(archivo, archivo_fbl5n)
            if isinstance(resultado, (list, tuple, dict)):
                total_exportados += len(resultado if not isinstance(resultado, dict) else resultado.values())
            else:
                total_exportados += 1
        messagebox.showinfo("Éxito", f"{total_exportados} archivo(s) generado(s) para {cliente}.")
    except Exception as ex:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error", f"❌ Error: {ex}")

# Interfaz principal
root = tk.Tk()
root.title("Procesador de Clientes")
root.geometry("500x600")
root.configure(bg="white")

# País
pais_var = tk.StringVar()
tk.Label(root, text="Selecciona el país", bg="white", font=("Arial", 14)).pack()
pais_dropdown = ttk.Combobox(root, textvariable=pais_var, values=paises, state="readonly")
pais_dropdown.pack(pady=5)
pais_dropdown.bind("<<ComboboxSelected>>", actualizar_clientes)

# Cliente
cliente_var = tk.StringVar()
cliente_dropdown = ttk.Combobox(root, textvariable=cliente_var, state="readonly")
cliente_dropdown.pack(pady=5)
cliente_dropdown.bind("<<ComboboxSelected>>", actualizar_visibilidad_fbl5n)

# Botones de carga
remittance_button = tk.Button(root, text="Adjuntar Remittance", command=seleccionar_remittance)
remittance_button.pack(pady=5)
remittance_status = tk.Label(root, text="Haz clic para adjuntar los Remittance", bg="white")
remittance_status.pack()

fbl5n_button = tk.Button(root, text="Adjuntar FBL5N", command=seleccionar_fbl5n)
fbl5n_button.pack(pady=5)
fbl5n_status = tk.Label(root, text="Haz clic para adjuntar el FBL5N", bg="white")
fbl5n_status.pack()

# Procesar
tk.Button(root, text="Generar Template", command=procesar_cliente, bg="#261E97", fg="white").pack(pady=20)

root.mainloop()