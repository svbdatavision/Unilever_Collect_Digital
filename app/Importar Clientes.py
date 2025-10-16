
import tkinter as tk
from tkinter import ttk, messagebox
from clientes.Colombia import olimpica, farmatodo, D1, copi, Cencosud

# Registro de clientes
CLIENTES = {
    "Olimpica": olimpica.procesar,
    "Farmatodo": farmatodo.procesar,
    "D1": D1.procesar,
    "Copidrogas": copi.procesar,
    "Cencosud": Cencosud.procesar
    #"Euro" : Euro.procesar
}

# Función que se ejecuta al hacer clic en el botón
def procesar_cliente():
    cliente = combo.get()
    try:
        procesar = CLIENTES[cliente]
        df = procesar()
        messagebox.showinfo("Éxito", f"✅ Proceso completado para {cliente}")
        print(df.head())  # Puedes redirigir esto a un archivo o ventana si lo deseas
    except Exception as e:
        import traceback
        traceback.print_exc()
        messagebox.showerror("Error", f"Ocurrió un error: \n{e}")

# Crear ventana principal
ventana = tk.Tk()
ventana.title("Importar Clientes")
ventana.geometry("300x150")

# Dropdown (Combobox)
label = tk.Label(ventana, text="Cliente:")
label.pack(pady=5)

combo = ttk.Combobox(ventana, values=list(CLIENTES.keys()))
combo.current(0)
combo.pack(pady=5)

# Botón
boton = tk.Button(ventana, text="Procesar", command=procesar_cliente)
boton.pack(pady=10)

# Ejecutar la interfaz
ventana.mainloop()
