import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from PIL import ImageTk, Image
import os

from Peru import (
    Template_Cencosud,
    Template_mayorsa,
    Template_NF,
    Template_SPSA,
    Template_Tottus
)

from Colombia import (
    Cencosud,
    copi,
    D1,
    farmatodo,
    olimpica,
    Euro,
    oxxo,
    Cruz_Verde
)
from Ecuador import (
    Favorita, 
    FARCOMED,
    DIFARE, 
    Mega_Santa_Maria, 
    FARMAENLACE,
    El_Rosado
)

archivos_remittance = []
archivo_fbl5n = None
pais_actual = None

paises = [
    "Peru",
    "Colombia",
    "Ecuador"
]

clientes_por_pais = {
    "Peru": [
        "Cencosud Peru",
        "Mayorsa",
        "Nortfarma",
        "SPSA",
        "Tottus"
    ],
    "Colombia": [
        "Cencosud Colombia",
        "Copidrogas",
        "D1",
        "Farmatodo",
        "Olimpica",
        "Euro",
        "Oxxo",
        "Cruz Verde"
    ],
    "Ecuador": [
        "Favorita", 
        "FARCOMED", 
        "DIFARE", 
        "Mega Santa Maria", 
        "Farmaenlace",
        "El Rosado"
    ]
}

procesadores = {
    # Peru
    "Cencosud Peru": Template_Cencosud,
    "Mayorsa": Template_mayorsa,
    "Nortfarma": Template_NF,
    "SPSA": Template_SPSA,
    "Tottus": Template_Tottus,
    # Colombia
    "Cencosud Colombia": Cencosud,
    "Copidrogas": copi,
    "D1": D1,
    "Farmatodo": farmatodo,
    "Olimpica": olimpica,
    "Euro": Euro,
    "Oxxo": oxxo,
    "Cruz Verde": Cruz_Verde,
    # Ecuador
    "Favorita": Favorita,
    "FARCOMED": FARCOMED,
    "DIFARE": DIFARE,
    "Mega Santa Maria": Mega_Santa_Maria,
    "Farmaenlace": FARMAENLACE,
    "El Rosado": El_Rosado
}

def seleccionar_remittance():
    global archivos_remittance
    archivos_remittance = filedialog.askopenfilenames(title="Select Remittance Files")
    remittance_status.config(text=f"{len(archivos_remittance)} remittance file(s) uploaded")

def seleccionar_fbl5n():
    global archivo_fbl5n
    archivo_fbl5n = filedialog.askopenfilename(title="Select FBL5N File")
    fbl5n_status.config(text="FBL5N File Uploaded")

def actualizar_clientes(event=None):
    global pais_actual
    pais_actual = pais_var.get()
    import config
    config.pais_actual = pais_actual  # ← Actualiza el módulo config
    clientes = clientes_por_pais.get(pais_actual, [])
    cliente_dropdown['values'] = clientes
    cliente_dropdown.set('')
    fbl5n_status.config(text="Click to attach the FBL5N")
    remittance_status.config(text="Click to attach the Remittance files")

def actualizar_visibilidad_fbl5n(event=None):
    cliente = cliente_var.get()
    if cliente in []:
        fbl5n_button.config(state="disabled")
    else:
        fbl5n_button.config(state="normal")

def procesar_cliente():
    cliente = cliente_var.get()
    if not archivos_remittance:
        messagebox.showwarning("Warning", "Please attach at least one Remittance file.")
        return
    if cliente not in [] and not archivo_fbl5n:
        messagebox.showwarning("Warning", "This client also requires the FBL5N file.")
        return
    try:
        total_exportados = 0
        for archivo in archivos_remittance:
            resultado = procesadores[cliente].procesar(archivo, archivo_fbl5n)
            if isinstance(resultado, (list, tuple, dict)):
                total_exportados += len(resultado if not isinstance(resultado, dict) else resultado.values())
            else:
                total_exportados += 1
        messagebox.showinfo("Success", f"{total_exportados} file(s) generated for {cliente}.")
    except Exception as ex:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error", f"❌ Error: {ex}")

# Interfaz principal
root = tk.Tk()
root.title("HighRadius Payment Template")
root.geometry("500x600")
root.configure(bg="white")

# País
pais_var = tk.StringVar()
tk.Label(root, text="Select the country", bg="white", font=("Arial", 14)).pack()
pais_dropdown = ttk.Combobox(root, textvariable=pais_var, values=paises, state="readonly")
pais_dropdown.pack(pady=5)
pais_dropdown.bind("<<ComboboxSelected>>", actualizar_clientes)

# Cliente
cliente_var = tk.StringVar()
cliente_dropdown = ttk.Combobox(root, textvariable=cliente_var, state="readonly")
cliente_dropdown.pack(pady=5)
cliente_dropdown.bind("<<ComboboxSelected>>", actualizar_visibilidad_fbl5n)

# Botones de carga
remittance_button = tk.Button(root, text="Attach Remittance", command=seleccionar_remittance)
remittance_button.pack(pady=5)
remittance_status = tk.Label(root, text="Click to attach the Remittance files", bg="white")
remittance_status.pack()

fbl5n_button = tk.Button(root, text="Attach FBL5N", command=seleccionar_fbl5n)
fbl5n_button.pack(pady=5)
fbl5n_status = tk.Label(root, text="Click to attach the FBL5N", bg="white")
fbl5n_status.pack()

# Procesar
tk.Button(root, text="Generate Template", command=procesar_cliente, bg="#261E97", fg="white").pack(pady=20)

root.mainloop()