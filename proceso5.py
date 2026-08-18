import gc
import win32com.client as win32

"""
Este archivo en python busca automatizar el proceso 5
Consiste en copiar el formato de la información de los proyectos del último mes
pegarlos y cambiarles el mes y otros datos.
Hay que tener cuidado en ejecutarlo una única vez al mes para no actualizarse de más.
"""

def actualizarmes():
    excel = None
    wb_archivo = None

    try:
        print("(1) Empieza el proceso")
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.Interactive = False
        excel.ScreenUpdating = False  # Optimización de rendimiento

        # ------------------------------------
        # --------- RUTAS Y ARCHIVOS ---------
        # ------------------------------------
        base_path = r"C:\Users\30630\OneDrive - Grupo Fahnos\Documentos\autom-pruebas\TDIACCM 2026 - copia.xlsm"
        wb_archivo = excel.Workbooks.Open(base_path)
        ws = wb_archivo.Worksheets("Base")

        xlUp = -4162

        # --- Última fila con datos en columna A ---
        ultima_fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        if ultima_fila <= 1:
            raise ValueError(
                "No se encontraron datos suficientes en la hoja 'Base'."
            )

        valor_a_uf = ws.Cells(ultima_fila, "A").Value
        mes_uf = ws.Cells(ultima_fila, "E").Value
        año_uf = ws.Cells(ultima_fila, "D").Value

        # --- Buscar la primera fila del bloque (mismo valor en columna A) ---
        primera_fila_rango = ultima_fila
        while (
            primera_fila_rango > 1
            and ws.Cells(primera_fila_rango - 1, "A").Value == valor_a_uf
        ):
            primera_fila_rango -= 1

        num_filas = ultima_fila - primera_fila_rango + 1
        fila_destino_inicio = ultima_fila + 1
        fila_destino_fin = ultima_fila + num_filas

        print(
            f"Bloque origen: filas {primera_fila_rango}-{ultima_fila} ({num_filas} filas)"
        )
        print(
            f"Bloque destino: filas {fila_destino_inicio}-{fila_destino_fin}"
        )

        # --- Copiar el bloque completo A:CG (mantiene formatos y fórmulas) ---
        rango_origen = ws.Range(f"A{primera_fila_rango}:CG{ultima_fila}")
        rango_destino = ws.Range(f"A{fila_destino_inicio}:CG{fila_destino_fin}")
        rango_origen.Copy(rango_destino)

        # --- Actualizar columna A (correlativo +1) ---
        nuevo_valor_a = int(valor_a_uf) + 1
        ws.Range(
            f"A{fila_destino_inicio}:A{fila_destino_fin}"
        ).Value = nuevo_valor_a

        # --- Actualizar columna E (mes siguiente) ---
        meses = [
            "Ene",
            "Feb",
            "Mar",
            "Abr",
            "May",
            "Jun",
            "Jul",
            "Ago",
            "Sep",
            "Oct",
            "Nov",
            "Dic",
        ]
        mes_uf_str = str(mes_uf).strip()

        idx = next(
            (
                i
                for i, m in enumerate(meses)
                if m.lower() == mes_uf_str.lower()
            ),
            None,
        )
        if idx is None:
            raise ValueError(
                f"Mes no reconocido en columna E: '{mes_uf_str}'"
            )

        nuevo_idx = (idx + 1) % 12
        nuevo_mes = meses[nuevo_idx]
        ws.Range(
            f"E{fila_destino_inicio}:E{fila_destino_fin}"
        ).Value = nuevo_mes

        # --- Si pasamos de Dic a Ene, incrementar el año en columna D ---
        if idx == 11:
            nuevo_año = int(año_uf) + 1
            ws.Range(
                f"D{fila_destino_inicio}:D{fila_destino_fin}"
            ).Value = nuevo_año
            print(f"Cambio de año detectado: {año_uf} -> {nuevo_año}")

        wb_archivo.Save()
        wb_archivo.Close()
        print("¡Proceso completado con éxito!")

    except Exception as e:
        print(f"Ocurrió un error durante la ejecución: {e}")
        if wb_archivo:
            try:
                wb_archivo.Close(SaveChanges=False)
            except Exception:
                pass

    finally:
        if excel:
            try:
                excel.ScreenUpdating = True
                excel.Interactive = True
                excel.Quit()
            except Exception:
                pass
        gc.collect()


# --- ESTA ES LA LÍNEA QUE FALTABA PARA EJECUTAR EL CÓDIGO ---
if __name__ == "__main__":
    actualizarmes()