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
    oxxo
)
from Ecuador import (
    El_Rosado,
    Favorita, 
    FARCOMED,
    DIFARE, 
    Mega_Santa_Maria, 
    FARMAENLACE
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
        "Oxxo"
    ],
    "Ecuador": [
        "Favorita", 
        "FARCOMED", 
        "DIFARE", 
        "Mega Santa Maria", 
        "Farmaenlace"
    ]
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

def main(page: ft.Page):
    page.title = "HighRadius Payment Template"
    page.bgcolor = "#FFFFFF"
    page.window_width = 500
    page.window_height = 600
    page.scroll = "auto"

    logo = ft.Image(src="unilever_logo.png", width=300, height=300)

    cliente_dropdown = ft.Dropdown(
        label="Select customer",
        width=350,
        options=[],
        bgcolor="#FFFFFF",
        border_color="#261E97",
        color="#FFFFFF"
    )
    
    pais_label = ft.Text("", size=18, weight="w600", color="#2E23C9")

    remittance_status = ft.Text("Click to attach the Remittances", text_align="center", color="black")
    fbl5n_status = ft.Text("Click to attach the FBL5N", text_align="center", color="black")

    remittance_uploader = ft.FilePicker(on_result=lambda e: handle_remittance(e))
    fbl5n_uploader = ft.FilePicker(on_result=lambda e: handle_fbl5n(e))
    page.overlay.extend([remittance_uploader, fbl5n_uploader])

    remittance_drop_area = ft.Container(
        content=ft.GestureDetector(
            content=ft.Column([
                ft.Icon(name="upload_file", size=40, color="#261E97"),
                remittance_status
            ], alignment="center", horizontal_alignment="center"),
            on_tap=lambda _: remittance_uploader.pick_files(allow_multiple=True)
        ),
        width=200,
        height=120,
        bgcolor="#F5F5F5",
        border=ft.border.all(2, "#261E97"),
        border_radius=10,
        padding=10,
        alignment=ft.alignment.center
    )

    fbl5n_drop_area = ft.Container(
        content=ft.GestureDetector(
            content=ft.Column([
                ft.Icon(name="upload_file", size=40, color="#261E97"),
                fbl5n_status
            ], alignment="center", horizontal_alignment="center"),
            on_tap=lambda _: fbl5n_uploader.pick_files()
        ),
        width=200,
        height=120,
        bgcolor="#F5F5F5",
        border=ft.border.all(2, "#261E97"),
        border_radius=10,
        padding=10,
        alignment=ft.alignment.center,
        visible=False
    )

    def handle_remittance(e):
        global archivos_remittance
        archivos_remittance = e.files
        remittance_status.value = f"{len(archivos_remittance)} Remittance file(s) uploaded"
        remittance_drop_area.bgcolor = "#DFF5E1"
        page.update()

    def handle_fbl5n(e):
        global archivo_fbl5n
        archivo_fbl5n = e.files[0].path if e.files else None
        fbl5n_status.value = "FBL5N file uploaded"
        fbl5n_drop_area.bgcolor = "#DFF5E1"
        page.update()

    def limpiar_campos():
        global archivos_remittance, archivo_fbl5n
        archivos_remittance = []
        archivo_fbl5n = None
        remittance_status.value = "Click to attach the remittance files"
        remittance_drop_area.bgcolor = "#FFFFFF"
        fbl5n_status.value = "Click to attach the FBL5N file"
        fbl5n_drop_area.bgcolor = "#FFFFFF"
        page.update()

    def actualizar_visibilidad_fbl5n(e):
        limpiar_campos()
        cliente = cliente_dropdown.value
        fbl5n_drop_area.visible = cliente not in ["Cencosud Peru", "Nortfarma", "SPSA", "Tottus"]
        page.update()

    def mostrar_procesamiento(pais):
        global pais_actual
        pais_actual = pais
        logo.width = 160
        logo.height = 160
        clientes = clientes_por_pais.get(pais, [])
        if not clientes:
            page.dialog = ft.AlertDialog(title=ft.Text("⚠️ No customers available for this country."))
            page.dialog.open = True
            page.update()
            return
        pais_label.value = pais
        cliente_dropdown.options = [ft.dropdown.Option(c) for c in clientes]
        cliente_dropdown.value = None
        cliente_dropdown.on_change = actualizar_visibilidad_fbl5n
        processing_view.visible = True
        home_view.visible = False
        page.update()

    def volver_inicio(e):
        logo.width = 400
        logo.height = 400
        limpiar_campos()
        processing_view.visible = False
        home_view.visible = True
        page.update()

    def procesar_cliente(e):
        cliente = cliente_dropdown.value
        clientes_solo_remittance = [
            "Cencosud Peru", 
            "Nortfarma", 
            "SPSA", 
            "Tottus"
        ]

        if not archivos_remittance:
            page.dialog = ft.AlertDialog(title=ft.Text("⚠️ Please attach at least one remittance file."))
            page.dialog.open = True
            page.update()
            return

        if cliente not in clientes_solo_remittance and not archivo_fbl5n:
            page.dialog = ft.AlertDialog(title=ft.Text("⚠️ This customer also requires the FBL5N file."))
            page.dialog.open = True
            page.update()
            return

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
            "DIFARE": DIFARE
        }

        try:
            if cliente in procesadores:
                total_exportados = 0

                for archivo in archivos_remittance:
                    resultado = procesadores[cliente].procesar(archivo.path, archivo_fbl5n)

                    # Soporte para múltiples salidas
                    if isinstance(resultado, (list, tuple, dict)):
                        for i, _ in enumerate(
                            resultado if not isinstance(resultado, dict) else resultado.values(), start=1
                        ):
                            total_exportados += 1
                            print(f"✅ File exported #{total_exportados} for {cliente}")
                    else:
                        total_exportados += 1
                        print(f"✅ File exported #{total_exportados} for {cliente}")

                page.dialog = ft.AlertDialog(
                    title=ft.Text(f"✅ {total_exportados} file(s) generated for {cliente}.")
                )
                page.dialog.open = True
                page.update()

            else:
                page.dialog = ft.AlertDialog(title=ft.Text("⚠️ Please select a valid customer."))
                page.dialog.open = True
                page.update()

        except Exception as ex:
            import traceback
            traceback.print_exc()
            print(f"❌ Error processing {cliente}: {ex}")
            page.dialog = ft.AlertDialog(title=ft.Text(f"❌ Error: {ex}"))
            page.dialog.open = True
            page.update()

    home_view = ft.Column(
        controls=[
            logo,
            ft.Text("Select the country", size=20, color="#000000"),
            ft.Row([
                ft.ElevatedButton(
                    text=pais,
                    on_click=lambda e, p=pais: mostrar_procesamiento(p),
                    width=150,
                    height=50,
                    bgcolor="#FFFFFF",
                    color="#2E23C9",
                    style=ft.ButtonStyle(
                        shape=ft.RoundedRectangleBorder(radius=8),
                        side=ft.BorderSide(width=2, color="#261E97")
                    )
                ) for pais in paises
            ], alignment="center", spacing=20)
        ],
        alignment="center",
        horizontal_alignment="center"
    )

    processing_view = ft.Column(
        controls=[
            ft.Row([ft.ElevatedButton(
                "← Back",
                on_click=volver_inicio,
                bgcolor="#261E97",
                color="white",
                style=ft.ButtonStyle(
                    shape=ft.RoundedRectangleBorder(radius=8)
                )
            )
            ], alignment="start"),
            logo,
            pais_label,
            cliente_dropdown,
            ft.Row(
                controls=[remittance_drop_area, fbl5n_drop_area],
                alignment="center",
                spacing=20
            ),
            ft.Row([
                ft.ElevatedButton("Generate Template", on_click=procesar_cliente, bgcolor="#261E97", color="white"),
                ft.ElevatedButton("Download FBL5N", bgcolor="white", color="#137023",
                                  style=ft.ButtonStyle(
                                      shape=ft.RoundedRectangleBorder(radius=8),
                                      side=ft.BorderSide(width=2, color="#137023")
                                  ))
            ], alignment="center", spacing=20)
        ],
        visible=False,
        alignment="start",
        horizontal_alignment="center"
    )

    page.add(home_view, processing_view)

ft.app(target=main)
