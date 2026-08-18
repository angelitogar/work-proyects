import os
import sys
import tkinter as tk
from tkinter import simpledialog, messagebox
import win32com.client as win32

"""
 Este archivo de python automatiza el proceso de añadir cuentas nuevas al proceso 1
 Practicamente es escribir los datos relativos a la nueva cuenta 
 Posteriormente busca una carpeta con su nombre y copia los vínculos de excel y los pega en el documento correspondiente.
 Hay que tener atención con los nombres del proyecto y su carpeta, además de los archivos de Excel Generados, que no tengan errorer
 """


# ==========================================
# 1. MOSTRAR FORMULARIO / VENTANA EMERGENTE
# ==========================================
root = tk.Tk()
root.withdraw()  # Oculta la ventana principal de Tkinter


# Aquí se le pide al usuario poner el nombe EXACTO de la nueva cuenta en una ventana emergente
nombre_cuenta = simpledialog.askstring(
    title="Agregar Nueva Cuenta",
    prompt="Ingresa el identificador de la nueva cuenta (ej. BANXICO 14):"
)

if not nombre_cuenta:
    messagebox.showwarning("Cancelado", "No se ingresó ninguna cuenta. Proceso detenido.")
    sys.exit()

nombre_cuenta = nombre_cuenta.strip()

# ==========================================
# 2. RUTAS Y BÚSQUEDA DE CARPETA Y ARCHIVO
# ==========================================
RUTA_BASE_CUENTAS = r"C:\Users\30630\OneDrive - Grupo Fahnos\Documentos\autom-pruebas"
RUTA_CONSOLIDADO = r"C:\Users\30630\OneDrive - Grupo Fahnos\Documentos\autom-pruebas\CONSOLIDADO IND 2026.xlsx"

carpeta_encontrada = None
archivo_encontrado = None

# Buscar carpeta que termine en el nombre de la cuenta
carpetas_coincidentes = []
for elemento in os.listdir(RUTA_BASE_CUENTAS):
    ruta_elemento = os.path.join(RUTA_BASE_CUENTAS, elemento)
    if os.path.isdir(ruta_elemento) and elemento.endswith(nombre_cuenta):
        carpetas_coincidentes.append(ruta_elemento)

if not carpetas_coincidentes:
    messagebox.showerror("Error", f"No se encontró ninguna carpeta que termine en '{nombre_cuenta}'.")
    sys.exit()

if len(carpetas_coincidentes) > 1:
    messagebox.showerror(
        "Error",
        f"Se encontró más de una carpeta que termina en '{nombre_cuenta}':\n\n"
        + "\n".join(carpetas_coincidentes)
        + "\n\nRevisa manualmente cuál usar."
    )
    sys.exit()

carpeta_encontrada = carpetas_coincidentes[0]

# Buscar archivo .xlsx que termine en {nombre_cuenta}.xlsx dentro de esa carpeta
for archivo in os.listdir(carpeta_encontrada):
    if archivo.endswith(f"{nombre_cuenta}.xlsx") and not archivo.startswith("~$"):
        archivo_encontrado = os.path.join(carpeta_encontrada, archivo)
        break

if not archivo_encontrado:
    messagebox.showerror("Error", f"No se encontró ningún archivo que termine en '{nombre_cuenta}.xlsx'.")
    sys.exit()

# ==========================================
# 3. FUNCIONES AUXILIARES Y CONFIGURACIÓN
# ==========================================
def obtener_hoja(workbook, opciones_nombre):
    """Busca una hoja probando variaciones de nombre."""
    if isinstance(opciones_nombre, str):
        opciones_nombre = [opciones_nombre]
    for nombre in opciones_nombre:
        try:
            return workbook.Worksheets(nombre)
        except Exception:
            continue
    return None


def obtener_ultima_fila(hoja, columna="B"):
    """
    Devuelve la fila siguiente a la última con un identificador de cuenta real.
    """
    fila = hoja.Cells(hoja.Rows.Count, columna).End(xlUp).Row
    return fila + 1


def construir_formula_vinculo(hoja_origen, fila, columna):
    """
    Construye directamente el texto de una fórmula de vínculo externo hacia
    una celda de la hoja de origen, garantizando un vínculo directo.
    """
    nombre_archivo = os.path.basename(hoja_origen.Parent.FullName)
    nombre_hoja = hoja_origen.Name
    direccion_celda = hoja_origen.Cells(fila, columna).Address  # ej. "$AG$41"
    return f"='[{nombre_archivo}]{nombre_hoja}'!{direccion_celda}"


# Mapeo de pestañas
PAREJAS_HOJAS = [
    (["ENTREGABLES"], ["ENTREGABLES"], 1, False),
    (["OBLIGACIONES"], ["OBLIGACIONES"], 1, False),
    (["INVENTARIO", "INVENTARIOS"], ["INVENTARIO", "INVENTARIOS"], 0, True)
]

MAPA_MESES_COLUMNAS = {
    "enero": "D", "ene": "D",
    "febrero": "E", "feb": "E",
    "marzo": "F", "mar": "F",
    "abril": "G", "abr": "G",
    "mayo": "H", "may": "H",
    "junio": "I", "jun": "I",
    "julio": "J", "jul": "J",
    "agosto": "K", "ago": "K",
    "septiembre": "L", "sep": "L", "sept": "L",
    "octubre": "M", "oct": "M",
    "noviembre": "N", "nov": "N",
    "diciembre": "O", "dic": "O"
}

# Constantes de Excel
xlUp = -4162
xlPasteFormats = -4122
xlValues = -4163
xlPart = 2

# ==========================================
# 4. ALGORITMO DE BÚSQUEDA Y VINCULACIÓN
# ==========================================
excel = None
wb_destino = None
wb_origen = None
cambios_guardados = False

try:
    print("Abriendo archivos...")
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb_origen = excel.Workbooks.Open(archivo_encontrado, UpdateLinks=False)
    wb_destino = excel.Workbooks.Open(RUTA_CONSOLIDADO, UpdateLinks=False)

    resumen_proceso = []

    for nombres_origen, nombres_destino, offset_col_origen, escribir_columna_a in PAREJAS_HOJAS:
        hoja_origen = obtener_hoja(wb_origen, nombres_origen)
        hoja_destino = obtener_hoja(wb_destino, nombres_destino)

        if hoja_origen is None or hoja_destino is None:
            resumen_proceso.append(f"• Hoja no encontrada en origen o destino ({nombres_origen[0]}). Omisión.")
            continue

        # --- A) OBTENER ÚLTIMA FILA LIBRE ---
        columna_nombre_cuenta = "A" if escribir_columna_a else "B"
        ultima_fila = obtener_ultima_fila(hoja_destino, columna_nombre_cuenta)

        wb_destino.Activate()
        hoja_destino.Activate()

        # 1. Copiar formato de la fila anterior (A a O)
        if ultima_fila > 2:
            hoja_destino.Range(f"A{ultima_fila - 1}:O{ultima_fila - 1}").Copy()
            hoja_destino.Range(f"A{ultima_fila}:O{ultima_fila}").PasteSpecial(Paste=xlPasteFormats)
            excel.CutCopyMode = False

        # 2. Asignar el nombre de la cuenta en la Columna B (o A si es INVENTARIOS)
        hoja_destino.Range(f"{columna_nombre_cuenta}{ultima_fila}").Value = nombre_cuenta

        # --- B) BUSCAR CELDA CON "2026" EN ORIGEN ---
        celda_2026 = hoja_origen.UsedRange.Find(
            What="2026", LookIn=xlValues, LookAt=xlPart, MatchCase=False
        )
        if celda_2026 is None:
            celda_2026 = hoja_origen.UsedRange.Find(
                What=2026, LookIn=xlValues, LookAt=xlPart, MatchCase=False
            )

        if celda_2026 is None:
            resumen_proceso.append(f"• {hoja_destino.Name}: No se encontró '2026' en el origen.")
            continue

        fila_2026 = celda_2026.Row
        col_2026 = celda_2026.Column
        fila_meses = fila_2026 + 2  # 2 celdas abajo del año 2026

        # --- C) ESCANEAR FILA DE MESES Y VINCULAR ---
        meses_vinculados = 0

        for c in range(col_2026, col_2026 + 50):
            val = hoja_origen.Cells(fila_meses, c).Value
            if val is None:
                continue

            texto_mes = str(val).strip().lower()

            if texto_mes in MAPA_MESES_COLUMNAS:
                col_destino = MAPA_MESES_COLUMNAS[texto_mes]

                fila_dato_origen = fila_meses - 1
                col_dato_origen = c + offset_col_origen

                # --- ASIGNACIÓN DIRECTA DE FÓRMULA DE VÍNCULO (SIN PORTAPAPELES) ---
                formula_vinculo = construir_formula_vinculo(hoja_origen, fila_dato_origen, col_dato_origen)
                hoja_destino.Range(f"{col_destino}{ultima_fila}").Formula = formula_vinculo

                meses_vinculados += 1

        resumen_proceso.append(f"• {hoja_destino.Name}: {meses_vinculados} meses vinculados en fila {ultima_fila}.")

    # Guardar cambios en el archivo CONSOLIDADO
    wb_destino.Save()
    cambios_guardados = True
    wb_origen.Close(SaveChanges=False)
    wb_destino.Close(SaveChanges=False)

    mensaje_final = f"Proceso completado para la cuenta '{nombre_cuenta}':\n\n" + "\n".join(resumen_proceso)
    messagebox.showinfo("¡Éxito!", mensaje_final)

except Exception as e:
    detalle = (
        f"Ocurrió un error: {e}\n\n"
        f"Cambios guardados en CONSOLIDADO antes del error: {'Sí' if cambios_guardados else 'No'}"
    )
    messagebox.showerror("Error durante el proceso", detalle)

finally:
    try:
        if wb_origen is not None:
            wb_origen.Close(SaveChanges=False)
    except Exception:
        pass
    try:
        if wb_destino is not None and not cambios_guardados:
            wb_destino.Close(SaveChanges=False)
    except Exception:
        pass
    if excel is not None:
        excel.Quit()