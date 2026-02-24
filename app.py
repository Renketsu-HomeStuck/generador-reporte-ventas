import os
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkcalendar import DateEntry
from datetime import datetime, date
from PIL import Image
import reporte


# ===== FUNCIÓN PARA RECURSOS (FUNCIONA EN .EXE) =====
def ruta_recurso(nombre_archivo):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, nombre_archivo)
    return os.path.join(os.path.abspath("."), nombre_archivo)


# ===== CONFIG GLOBAL =====
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


# ===== VENTANA =====
ventana = ctk.CTk()
ventana.title("Generador de Reporte de Ventas")
ventana.geometry("580x680")
ventana.resizable(False, False)

try:
    ventana.iconbitmap("icono.ico")
except Exception:
    pass


# ===== HEADER CON LOGO =====
header_frame = ctk.CTkFrame(ventana, corner_radius=0)
header_frame.pack(fill="x")

logo_path = ruta_recurso("logo.png")

logo_img = ctk.CTkImage(
    light_image=Image.open(logo_path),
    size=(90, 90)
)

logo_label = ctk.CTkLabel(
    header_frame,
    image=logo_img,
    text=""
)
logo_label.pack(pady=(15, 5))

titulo = ctk.CTkLabel(
    header_frame,
    text="Generador de Reporte de Ventas",
    font=("Segoe UI", 24, "bold")
)
titulo.pack(pady=(0, 15))


# ===== FUNCIONES =====
def seleccionar_archivo():
    archivo = filedialog.askopenfilename(
        filetypes=[("Archivos Excel", "*.xlsx")]
    )
    if archivo:
        entry_archivo.delete(0, "end")
        entry_archivo.insert(0, archivo)


def limpiar():
    entry_archivo.delete(0, "end")
    total_general_label.configure(text="Total General: -")
    producto_label.configure(text="Producto Más Vendido: -")
    dia_label.configure(text="Día con Mayor Venta: -")
    monto_label.configure(text="Monto Día Mayor Venta: -")
    estado_label.configure(text="Estado: Listo")
    progress.set(0)


def generar():
    archivo = entry_archivo.get().strip()
    fecha_inicio = date_inicio.get()
    fecha_fin = date_fin.get()

    if not archivo:
        messagebox.showerror("Error", "Selecciona un archivo.")
        return

    btn_generar.configure(text="Procesando...", state="disabled")
    estado_label.configure(text="Estado: Procesando...")
    progress.set(0.2)
    ventana.update()

    try:
        df = reporte.leer_datos(archivo)
        reporte.validar_columnas(df)

        df = reporte.aplicar_filtro(df, fecha_inicio, fecha_fin)
        progress.set(0.5)
        ventana.update()

        df, resumen, total_por_vendedor, ventas_por_dia = (
            reporte.calcular_metricas(df)
        )

        total_general = resumen.loc[
            resumen["Métrica"] == "Total General", "Valor"
        ].values[0]

        producto_mas_vendido = resumen.loc[
            resumen["Métrica"] == "Producto Más Vendido", "Valor"
        ].values[0]

        dia_mayor = resumen.loc[
            resumen["Métrica"] == "Día con Mayor Venta", "Valor"
        ].values[0]

        monto_dia = resumen.loc[
            resumen["Métrica"] == "Monto Día Mayor Venta", "Valor"
        ].values[0]

        total_general_label.configure(
            text=f"Total General: ${float(total_general):,.0f}"
        )
        producto_label.configure(
            text=f"Producto Más Vendido: {producto_mas_vendido}"
        )
        dia_label.configure(
            text=f"Día con Mayor Venta: {dia_mayor}"
        )
        monto_label.configure(
            text=f"Monto Día Mayor Venta: ${float(monto_dia):,.0f}"
        )

        progress.set(0.8)
        ventana.update()

        carpeta_reportes = os.path.join(
            os.path.dirname(archivo),
            "Reportes"
        )
        os.makedirs(carpeta_reportes, exist_ok=True)

        nombre_sugerido = (
            f"reporte_ventas_"
            f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        )

        ruta_guardado = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Archivos Excel", "*.xlsx")],
            initialfile=nombre_sugerido,
            initialdir=carpeta_reportes,
            title="Guardar reporte como"
        )

        if not ruta_guardado:
            estado_label.configure(text="Estado: Cancelado")
            return

        reporte.generar_reporte(
            ruta_guardado,
            df,
            resumen,
            total_por_vendedor,
            ventas_por_dia
        )

        progress.set(1)
        estado_label.configure(text="Estado: Reporte generado")

        os.startfile(ruta_guardado)

        messagebox.showinfo(
            "Éxito",
            f"Reporte generado correctamente:\n\n{ruta_guardado}"
        )

    except ValueError as e:
        if str(e) == "NO_DATA_RANGE":
            messagebox.showwarning(
                "Sin datos",
                "No hay ventas en el rango seleccionado.\n\n"
                "Revisa las fechas o amplía el período."
            )
        else:
            messagebox.showerror("Error", str(e))
        estado_label.configure(text="Estado: Error")

    except Exception as e:
        messagebox.showerror("Error", str(e))
        estado_label.configure(text="Estado: Error")

    finally:
        btn_generar.configure(
            text="Generar Reporte",
            state="normal"
        )
        ventana.after(800, lambda: progress.set(0))


# ===== SECCIÓN PRINCIPAL =====
main_frame = ctk.CTkFrame(ventana)
main_frame.pack(pady=15, padx=20, fill="both", expand=True)

entry_archivo = ctk.CTkEntry(
    main_frame,
    width=460,
    placeholder_text="Selecciona un archivo Excel..."
)
entry_archivo.pack(pady=10)

btn_archivo = ctk.CTkButton(
    main_frame,
    text="Seleccionar archivo",
    command=seleccionar_archivo,
    width=250
)
btn_archivo.pack(pady=5)

info_label = ctk.CTkLabel(
    main_frame,
    text="Formato requerido: fecha, vendedor, producto, cantidad, precio",
    font=("Segoe UI", 11),
    text_color="gray"
)
info_label.pack(pady=5)


# ===== FECHAS =====
frame_fechas = ctk.CTkFrame(main_frame)
frame_fechas.pack(pady=15)

ctk.CTkLabel(frame_fechas, text="Fecha Inicio:").grid(row=0, column=0, padx=10, pady=5)
date_inicio = DateEntry(frame_fechas, width=14, date_pattern="yyyy-mm-dd")
date_inicio.grid(row=0, column=1, padx=10, pady=5)

ctk.CTkLabel(frame_fechas, text="Fecha Fin:").grid(row=1, column=0, padx=10, pady=5)
date_fin = DateEntry(frame_fechas, width=14, date_pattern="yyyy-mm-dd")
date_fin.grid(row=1, column=1, padx=10, pady=5)

# ===== FECHAS INTELIGENTES (PRIMER DÍA DEL MES HASTA HOY) =====
hoy = date.today()
primer_dia_mes = hoy.replace(day=1)

date_inicio.set_date(primer_dia_mes)
date_fin.set_date(hoy)




# ===== BOTONES =====
btn_generar = ctk.CTkButton(
    main_frame,
    text="Generar Reporte",
    command=generar,
    width=300,
    height=45
)
btn_generar.pack(pady=10)

btn_limpiar = ctk.CTkButton(
    main_frame,
    text="Limpiar",
    command=limpiar,
    width=200
)
btn_limpiar.pack(pady=5)


estado_label = ctk.CTkLabel(
    main_frame,
    text="Estado: Listo",
    font=("Segoe UI", 12)
)
estado_label.pack(pady=5)

progress = ctk.CTkProgressBar(main_frame, width=350)
progress.pack(pady=5)
progress.set(0)


# ===== RESULTADOS =====
frame_resultados = ctk.CTkFrame(main_frame)
frame_resultados.pack(pady=15, fill="x")

total_general_label = ctk.CTkLabel(frame_resultados, text="Total General: -", font=("Segoe UI", 14))
total_general_label.pack(pady=5)

producto_label = ctk.CTkLabel(frame_resultados, text="Producto Más Vendido: -", font=("Segoe UI", 14))
producto_label.pack(pady=5)

dia_label = ctk.CTkLabel(frame_resultados, text="Día con Mayor Venta: -", font=("Segoe UI", 14))
dia_label.pack(pady=5)

monto_label = ctk.CTkLabel(frame_resultados, text="Monto Día Mayor Venta: -", font=("Segoe UI", 14))
monto_label.pack(pady=5)


# ===== FOOTER =====
footer = ctk.CTkLabel(
    ventana,
    text="v1.0  •  Rene Lavanchy",
    font=("Segoe UI", 10)
)
footer.pack(pady=10)


ventana.mainloop()