import os
import time
import win32com.client as win32

def ejecutar_con_reintento(func, reintentos=5, espera=1.5):
    """Ejecuta una función COM reintentando si Excel responde que está ocupado."""
    for i in range(reintentos):
        try:
            return func()
        except Exception as e:
            # Captura el error de "llamada rechazada" (-2147418111)
            if "-2147418111" in str(e) or "rechazada" in str(e).lower():
                time.sleep(espera)
            else:
                raise e
    return func()

def automatizar_excel():
    excel = None
    wb_principal = None
    
    try:
        print("Iniciando Excel en segundo plano...")
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AskToUpdateLinks = False
        excel.Interactive = False  # Previene que la interfaz bloquee las llamadas COM

        ruta_principal = r"C:\Users\30630\Grupo Fahnos\Agustin Eduardo Martinez Fuentes - IMX-DIR_IMX-ACC\ACCM DIR\FACTURACION\Facturacion Consolidada Actual -1.xlsx"
        archivos_destino = [
            r"C:\Users\30630\Grupo Fahnos\Agustin Eduardo Martinez Fuentes - IMX-DIR_IMX-ACC\ACCM DIR\FACTURACION\Reporte de Facturacion y Cobranza Flujo y Ppto v15 -1.xlsx",
            r"C:\Users\30630\Grupo Fahnos\Agustin Eduardo Martinez Fuentes - IMX-DIR_IMX-ACC\ACCM DIR\FACTURACION\Reporte de Facturacion y Cobranza Flujo y Ppto v16 -1.xlsx"
        ]


        print("Abriendo archivo principal...")
        wb_principal = excel.Workbooks.Open(ruta_principal)
        time.sleep(3)  # Pausa estratégica para permitir la carga de Power Query

        # 1. Configurar conexiones con reintentos inteligentes
        print("Configurando conexiones de Power Query...")
        for conn in wb_principal.Connections:
            def deshabilitar_background():
                if hasattr(conn, "OLEDBConnection"):
                    conn.OLEDBConnection.BackgroundQuery = False
                elif hasattr(conn, "ODBCConnection"):
                    conn.ODBCConnection.BackgroundQuery = False

            try:
                ejecutar_con_reintento(deshabilitar_background)
            except Exception:
                pass

        # 2. Actualizar Consultas y Tablas
        print("Actualizando Consultas y Tablas (RefreshAll)...")
        ejecutar_con_reintento(lambda: wb_principal.RefreshAll())
        
        # Breve pausa para dar tiempo a procesar los datos
        time.sleep(2)

        print("Recalculando fórmulas en todo el libro...")
        ejecutar_con_reintento(lambda: excel.CalculateFull())

        print("Guardando cambios en el archivo principal...")
        ejecutar_con_reintento(lambda: wb_principal.Save())

        # 3. Leer los datos de BaseFact a memoria
        print("Copiando datos de BaseFact...")
        ws_basefact = wb_principal.Worksheets("BaseFact")
        
        ultima_fila_fact = ws_basefact.Cells(ws_basefact.Rows.Count, "A").End(-4162).Row  # -4162 = xlUp

        if ultima_fila_fact >= 2:
            datos_origen = ws_basefact.Range(f"A2:Q{ultima_fila_fact}").Value

            for ruta_dest in archivos_destino:
                print(f"Pegando en: {os.path.basename(ruta_dest)}...")
                wb_destino = excel.Workbooks.Open(ruta_dest, UpdateLinks=0)
                ws_destino_fact = wb_destino.Worksheets("Base")

                # Limpiar datos anteriores desde A2 en adelante
                ultima_fila_destino = ws_destino_fact.Cells(ws_destino_fact.Rows.Count, "A").End(-4162).Row
                if ultima_fila_destino > 1:
                    ws_destino_fact.Range(f"A2:Q{ultima_fila_destino}").ClearContents()

                # Asignación directa de los valores en memoria
                ws_destino_fact.Range(f"A2:Q{ultima_fila_fact}").Value = datos_origen

                wb_destino.Save()
                wb_destino.Close()
        else:
            print("No hay datos suficientes para copiar en BaseFact.")

        wb_principal.Close()
        print("¡Proceso completado con éxito!")

    except Exception as e:
        print(f"Ocurrió un error durante la ejecución: {e}")

    finally:
        if excel:
            try:
                excel.CutCopyMode = False
                excel.Interactive = True  # Restaurar la interacción
                excel.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    automatizar_excel()