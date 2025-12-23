import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import config

# Importar codigos de clientes
from clientes.Colombia import (
    olimpica,
    farmatodo,
    D1,
    Cencosud,
    copi,
    Euro,
    oxxo,
    Cruz_Verde,
    Jeronimo,
    Prueba_Colombia_xlsx,
    Prueba_Colombia_pdf
)

from clientes.Ecuador import (
    El_Rosado,
    Prueba_Ecuador_xlsx,
    Prueba_Ecuador_pdf
)
from clientes.Peru import (
    Template_Cencosud,
    Template_NF,
    Template_SPSA,
    Template_Tottus,
    Prueba_Peru_xlsx,
    Prueba_Peru_pdf
)

# Registro de clientes por país
CLIENTES = {
    "Colombia": {
        "Olimpica": olimpica.procesar,
        "Farmatodo": farmatodo.procesar,
        "D1": D1.procesar,
        "Cencosud": Cencosud.procesar,
        "Copidrogas": copi.procesar,
        "Euro": Euro.procesar,
        "Oxxo": oxxo.procesar,
        "Cruz_Verde": Cruz_Verde.procesar,
        "Jeronimo": Jeronimo.procesar,
        "Prueba xlsx": Prueba_Colombia_xlsx.procesar,
        "Prueba pdf": Prueba_Colombia_pdf.procesar
    },
    "Ecuador": {
        "El Rosado": El_Rosado.procesar,
        "Prueba xlsx": Prueba_Ecuador_xlsx.procesar,
        "Prueba pdf": Prueba_Ecuador_pdf.procesar
    },
    "Peru": {
        "Cencosud Peru": Template_Cencosud.procesar,
        "NF": Template_NF.procesar,
        "SPSA": Template_SPSA.procesar,
        "Template_Tottus": Template_Tottus.procesar,
        "Prueba xlsx": Prueba_Peru_xlsx.procesar,
        "Prueba pdf": Prueba_Peru_pdf.procesar
    }
}

# =====================================================
# Función principal de procesamiento
# =====================================================
def procesar_cliente():
    pais = combo_pais.get()
    cliente = combo_cliente.get()

    # 👇 DEFINIR PAÍS UNA SOLA VEZ
    config.pais_actual = pais

    if not pais or not cliente:
        messagebox.showwarning("Advertencia", "Seleccione un país y un cliente antes de procesar.")
        return

    try:
        procesar = CLIENTES[pais][cliente]
        result = procesar()

        messagebox.showinfo("Éxito", f"✅ Proceso completado para {cliente} ({pais})")
        print(f"\n=== Resultados para {cliente} ({pais}) ===")

        if isinstance(result, dict):
            for key, df in result.items():
                print(f"\n📄 {key}:")
                print(df.head())
                print("-" * 50)

        elif isinstance(result, (list, tuple)):
            for i, df in enumerate(result, start=1):
                print(f"\n📄 DataFrame {i}:")
                print(df.head())
                print("-" * 50)

        elif isinstance(result, pd.DataFrame):
            print(result.head())

        else:
            print("⚠️ El resultado no es un DataFrame ni una lista/dict reconocida.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error", f"Ocurrió un error:\n{e}")

# Función para actualizar lista de clientes al cambiar país
def actualizar_clientes(event=None):
    pais = combo_pais.get()
    clientes_disponibles = list(CLIENTES[pais].keys())
    combo_cliente["values"] = clientes_disponibles
    if clientes_disponibles:
        combo_cliente.current(0)

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Importar Clientes")
ventana.geometry("340x200")

# Widgets: País
label_pais = tk.Label(ventana, text="País:")
label_pais.pack(pady=5)

combo_pais = ttk.Combobox(ventana, values=list(CLIENTES.keys()), state="readonly")
combo_pais.current(0)
combo_pais.pack(pady=5)
combo_pais.bind("<<ComboboxSelected>>", actualizar_clientes)

# Widgets: Cliente
label_cliente = tk.Label(ventana, text="Cliente:")
label_cliente.pack(pady=5)

combo_cliente = ttk.Combobox(ventana, state="readonly")
combo_cliente.pack(pady=5)
actualizar_clientes()  # Inicializar lista de clientes según país seleccionado por defecto

# Botón principal
boton = tk.Button(ventana, text="Procesar", command=procesar_cliente)
boton.pack(pady=10)

# Ejecutar la interfaz
ventana.mainloop()
