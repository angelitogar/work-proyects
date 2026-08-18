import os
import time
import shutil
import win32com.client as win32

"""
 Este archivo en python automatiza el proceso 2 (Facturación).
 Consiste en actualizar una base de datos y posteriormente recolectar datos para pegarlos en otra base de datos
 Se utiliza un diseño con archivos temporales para evitar errores donde no se soporte el tamaño del archivo y lo envie a otra ruta
 Hay que tener cuidado con el output y que se haya actualizado correctamente
"""

#Al ser un archivo muy grande, hay que reintentar en caso de que Excel se trabe o esté ocupado actualizandose
def ejecutar_con_reintento(func, reintentos=15, espera=2.0):
    """Ejecuta una función COM reintentando si Excel responde que está ocupado."""
    for i in range(reintentos):
        try:
            return func()
        except Exception as e:
            if i == reintentos - 1:
                raise e
            time.sleep(espera)

def automatizar_excel():
    excel = None
    wb_principal = None
    
    # 1. Definir directorio temporal local (totalmente aislado de OneDrive)
    # Esto es para evitar errores de sobreescritura y de carga del archivo
    temp_dir = os.path.join(os.environ["USERPROFILE"], "Temp_Automation")
    os.makedirs(temp_dir, exist_ok=True)

    # Rutas originales en SharePoint / OneDrive Sync
    ruta_principal_orig = r"C:\Users\30630\Grupo Fahnos\Agustin Eduardo Martinez Fuentes - IMX-DIR_IMX-ACC\ACCM DIR\FACTURACION\Facturacion Consolidada Actual -1.xlsx"
    archivos_destino_orig = [
        r"C:\Users\30630\Grupo Fahnos\Agustin Eduardo Martinez Fuentes - IMX-DIR_IMX-ACC\ACCM DIR\FACTURACION\Reporte de Facturacion y Cobranza Flujo y Ppto v15 -1.xlsx",
        r"C:\Users\30630\Grupo Fahnos\Agustin Eduardo Martinez Fuentes - IMX-DIR_IMX-ACC\ACCM DIR\FACTURACION\Reporte de Facturacion y Cobranza Flujo y Ppto v16 -1.xlsx"
    ]

    # Definir rutas temporales
    temp_principal = os.path.join(temp_dir, os.path.basename(ruta_principal_orig))
    archivos_temp_destino = []

    print("Copiando archivos a espacio de trabajo local...")
    # Copiar archivo principal a carpeta temporal
    shutil.copy2(ruta_principal_orig, temp_principal)

    # Copiar archivos destino a carpeta temporal
    for r_dest in archivos_destino_orig:
        t_dest = os.path.join(temp_dir, os.path.basename(r_dest))
        shutil.copy2(r_dest, t_dest)
        archivos_temp_destino.append((t_dest, r_dest))

    try:
        print("Iniciando Excel en segundo plano...")
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False

        print("Abriendo archivo principal desde carpeta local temporal...")
        wb_principal = ejecutar_con_reintento(lambda: excel.Workbooks.Open(temp_principal))
        time.sleep(2)

        # Configurar conexiones de Power Query (Protegido con reintento)
        print("Configurando conexiones de Power Query...")
        conexiones = ejecutar_con_reintento(lambda: [conn for conn in wb_principal.Connections])
        
        for conn in conexiones:
            def deshabilitar_background(c=conn):
                if hasattr(c, "OLEDBConnection"):
                    c.OLEDBConnection.BackgroundQuery = False
                elif hasattr(c, "ODBCConnection"):
                    c.ODBCConnection.BackgroundQuery = False

            try:
                ejecutar_con_reintento(deshabilitar_background)
            except Exception:
                pass

        # Actualizar Consultas y Tablas
        print("Actualizando Consultas y Tablas (RefreshAll)...")
        ejecutar_con_reintento(lambda: wb_principal.RefreshAll())
        
        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            time.sleep(3)

        print("Recalculando fórmulas...")
        ejecutar_con_reintento(lambda: excel.CalculateFull())

        # Actualizar la Tabla Dinámica
        print("Actualizando Tabla Dinámica 'Tabla dinámica1'...")
        try:
            ws_td = wb_principal.Worksheets("TD")
            pt = ws_td.PivotTables("Tabla dinámica1")
            ejecutar_con_reintento(lambda: pt.PivotCache().Refresh())
            ejecutar_con_reintento(lambda: pt.Update())
            print("Tabla dinámica1 actualizada con éxito.")
        except Exception as pt_err:
            print(f"Advertencia al actualizar 'Tabla dinámica1': {pt_err}")
            try:
                for pc in wb_principal.PivotCaches():
                    ejecutar_con_reintento(lambda: pc.Refresh())
            except Exception:
                pass

        ejecutar_con_reintento(lambda: excel.CalculateFull())

        # Comparar dos celdas de control antes de continuar
        ws_consolidado = wb_principal.Worksheets("Consolidado_")
        valor_1 = ws_consolidado.Range("Z1").Value
        valor_2 = ws_consolidado.Range("AA1").Value

        if valor_1 == "Vencimiento" and valor_2 == "Año-mes":
            print(f"Validación OK: A1 ({valor_1}) coincide con B1 ({valor_2}).")
        else:
            print(f"ERROR: A1 ({valor_1}) no coincide con B1 ({valor_2}). Deteniendo proceso.")
            return

        print("Guardando cambios en archivo principal temporal...")
        ejecutar_con_reintento(lambda: wb_principal.Save())

        # Copiar datos de BaseFact a memoria
        print("Copiando datos de BaseFact...")
        ws_basefact = wb_principal.Worksheets("BaseFact")
        ultima_fila_fact = ws_basefact.Cells(ws_basefact.Rows.Count, "A").End(-4162).Row

        if ultima_fila_fact >= 2:
            datos_origen = ws_basefact.Range(f"A5:Q{ultima_fila_fact - 1}").Value

            for temp_destino, _ in archivos_temp_destino:
                nombre_archivo = os.path.basename(temp_destino)
                print(f"Pegando en temporal: {nombre_archivo}...")
                
                wb_destino = ejecutar_con_reintento(lambda: excel.Workbooks.Open(temp_destino, UpdateLinks=0))

                ws_destino_fact = wb_destino.Worksheets("Base")
                ultima_fila_destino = ws_destino_fact.Cells(ws_destino_fact.Rows.Count, "A").End(-4162).Row
                
                if ultima_fila_destino > 1:
                    ws_destino_fact.Range(f"A2:Q{ultima_fila_destino}").ClearContents()

                ws_destino_fact.Range(f"A2:Q{ultima_fila_fact}").Value = datos_origen

                ejecutar_con_reintento(lambda: wb_destino.Save())
                ejecutar_con_reintento(lambda: wb_destino.Close(SaveChanges=False))
        else:
            print("No hay datos suficientes para copiar en BaseFact.")

        # Cerrar libro principal
        ejecutar_con_reintento(lambda: wb_principal.Close(SaveChanges=False))

    except Exception as e:
        print(f"Ocurrió un error durante el procesamiento en Excel: {e}")
        return

    finally:
        # Cerrar Excel por completo para LIBERAR los archivos antes de copiar
        if excel:
            try:
                excel.CutCopyMode = False
                excel.Quit()
                print("Excel cerrado correctamente.")
            except Exception:
                pass

    # =========================================================================
    # REEMPLAZO FINAL A SHAREPOINT/ONEDRIVE CON PYTHON
    # =========================================================================
    print("\nSincronizando cambios de vuelta a SharePoint/OneDrive...")
    time.sleep(2)  # Pausa de seguridad para liberación completa de handles en Windows

    try:
        # 1. Copiar archivo principal procesado
        print(f"Actualizando en nube: {os.path.basename(ruta_principal_orig)}")
        shutil.copy2(temp_principal, ruta_principal_orig)

        # 2. Copiar archivos destino procesados
        for temp_path, orig_path in archivos_temp_destino:
            print(f"Actualizando en nube: {os.path.basename(orig_path)}")
            shutil.copy2(temp_path, orig_path)

        # 3. Limpiar directorio temporal
        shutil.rmtree(temp_dir, ignore_errors=True)
        print("\n¡Proceso completado con éxito! Los archivos fueron actualizados y OneDrive iniciará la sincronización automáticamente.")

    except Exception as e:
        print(f"\nError al reemplazar los archivos finales: {e}")

if __name__ == "__main__":
    automatizar_excel()