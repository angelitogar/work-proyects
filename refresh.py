import os
import sys
import win32com.client as win32

def actualizar_y_guardar_excel(nombre_o_ruta_archivo):
    """
    Actualiza vínculos y fuentes de datos en un archivo de Excel y lo guarda.
    Funciona con rutas relativas (misma carpeta) o rutas absolutas.
    """
    # 1. Obtener la ruta absoluta completa del archivo
    if os.path.isabs(nombre_o_ruta_archivo):
        ruta_absoluta = nombre_o_ruta_archivo
    else:
        # Si es ruta relativa, la combina con la ubicación real de este archivo .py
        directorio_actual = os.path.dirname(os.path.abspath(__file__))
        ruta_absoluta = os.path.join(directorio_actual, nombre_o_ruta_archivo)
    
    # 2. Comprobar que el archivo realmente existe antes de abrir Excel
    if not os.path.exists(ruta_absoluta):
        print(f"[ERROR] No se encontró el archivo en: {ruta_absoluta}")
        return

    print(f"Abriendo archivo: {os.path.basename(ruta_absoluta)}...")
    
    excel = None
    try:
        # 3. Iniciar Excel en segundo plano
        excel = win32.Dispatch("Excel.Application")
        excel.Visible = True       # Mantiene la ventana de Excel oculta
        excel.DisplayAlerts = False  # Evita mensajes flotantes emergentes

        # UpdateLinks=3 actualiza automáticamente los vínculos a otros libros de Excel
        
        workbook = excel.Workbooks.Open(
            ruta_absoluta,
            UpdateLinks=3,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True
        )
               
        print("Actualizando datos, conexiones externas y recálculos...")
        workbook.RefreshAll()
        print("Conexiones detectadas y procesadas:")
        for conexion in workbook.Connections:
            print(f" - Conexión: {conexion.Name}")
        
        # Esperar a que terminen las consultas asíncronas (Power Query, SQL, etc.)
        excel.CalculateUntilAsyncQueriesDone()
        
        print("Guardando cambios...")
        workbook.Save()
        workbook.Close(SaveChanges=True)
        print("¡Proceso completado con éxito!")

    except Exception as e:
        print(f"[ERROR] Ocurrió un fallo durante el proceso: {e}")
        
    finally:
        # 4. Asegurar el cierre completo del proceso de Excel en la RAM
        if excel is not None:
            excel.Quit()

if __name__ == "__main__":
    # --- CONFIGURACIÓN ---
    # Si el Excel está en la MISMA carpeta que este script:
    #ARCHIVO_EXCEL = "CONSOLIDADO IND 2026.xlsx"

    # Si está en OTRA carpeta (Ruta absoluta):
    ARCHIVO_EXCEL = r"C:\Users\30630\Grupo Fahnos\Agustin Eduardo Martinez Fuentes - IMX-DIR_IMX-ACC\ACCM DIR\CONSOLIDADOS\CONSOLIDADO IND 2026.xlsx"

    actualizar_y_guardar_excel(ARCHIVO_EXCEL)