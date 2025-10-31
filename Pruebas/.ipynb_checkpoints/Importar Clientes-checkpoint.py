import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from clientes.Colombia import (
    olimpica,
    farmatodo,
    D1,
    Cencosud,
    Euro,
    El_Rosado,
    oxxo
)

# Registro de clientes
CLIENTES = {
    "Olimpica": olimpica.procesar,
    "Farmatodo": farmatodo.procesar,
    "D1": D1.procesar,
    "Copidrogas": copi.procesar,
    "Cencosud": Cencosud.procesar
    # "Euro": Euro.procesar
}

# Función principal
def procesar_cliente():
    cliente = combo.get()
    try:
        procesar = CLIENTES[cliente]
        result = procesar()

        messagebox.showinfo("Éxito", f"✅ Proceso completado para {cliente}")
        print(f"\n=== Resultados para {cliente} ===")

        # Si el resultado es un diccionario de DataFrames
        if isinstance(result, dict):
            for key, df in result.items():
                print(f"\n📄 {key}:")
                print(df.head())
                print("-" * 50)

        # Si el resultado es una lista o tupla de DataFrames
        elif isinstance(result, (list, tuple)):
            for i, df in enumerate(result, start=1):
                print(f"\n📄 DataFrame {i}:")
                print(df.head())
                print("-" * 50)

        # Si es un solo DataFrame
        elif isinstance(result, pd.DataFrame):
            print(result.head())

        else:
            print("⚠️ El resultado no es un DataFrame ni una lista/dict reconocida.")

    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error", f"Ocurrió un error: \n{e}")

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Importar Clientes")
ventana.geometry("320x180")

# Dropdown (Combobox)
label = tk.Label(ventana, text="Cliente:")
label.pack(pady=5)

combo = ttk.Combobox(ventana, values=list(CLIENTES.keys()), state="readonly")
combo.current(0)
combo.pack(pady=5)

# Botón
boton = tk.Button(ventana, text="Procesar", command=procesar_cliente)
boton.pack(pady=10)

# Ejecutar la interfaz
ventana.mainloop()
