
import os
import glob
from datetime import datetime

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CARPETA_ARCHIVOS = os.path.join(BASE_DIR, "archivos")
CARPETA_MODULES = os.path.join(BASE_DIR, "modules")
CARPETA_OUTPUTS = os.path.join(BASE_DIR, "outputs")
CARPETA_LOGS = os.path.join(BASE_DIR, "logs")

# Crear carpetas de salida si no existen
os.makedirs(CARPETA_OUTPUTS, exist_ok=True)
os.makedirs(CARPETA_LOGS, exist_ok=True)


# =========================================================
# ARCHIVOS DE ENTRADA
# =========================================================

# Buscar el archivo BO dentro de archivos
PATRON_BO = os.path.join(
    CARPETA_ARCHIVOS,
    "Reporte de Notas de Crédito - Casos*.xlsx"
)

ARCHIVOS_BO = glob.glob(PATRON_BO)

if not ARCHIVOS_BO:
    raise FileNotFoundError(
        "No se encontró el archivo BO en la carpeta archivos."
    )

# Si existen varios BO, seleccionar el más reciente
RUTA_BO = max(ARCHIVOS_BO, key=os.path.getmtime)

# Export de Siebel: se genera una ruta nueva para cada ejecución
def generar_ruta_export_siebel():

    marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    ruta = os.path.join(
        CARPETA_OUTPUTS,
        f"Export_Siebel_{marca_tiempo}.csv"
    )

    sufijo = 1
    while os.path.exists(ruta):
        ruta = os.path.join(
            CARPETA_OUTPUTS,
            f"Export_Siebel_{marca_tiempo}_{sufijo}.csv"
        )
        sufijo += 1

    return ruta


# =========================================================
# ARCHIVOS DE SALIDA
# =========================================================

RUTA_BO_RELLENADO = os.path.join(
    CARPETA_OUTPUTS,
    "Reporte_BO_Rellenado.xlsx"
)

RUTA_REGISTRO = os.path.join(
    CARPETA_LOGS,
    "Registro_Modificaciones.xlsx"
)


# =========================================================
# CONFIGURACIÓN DE EXCEL
# =========================================================

HOJA_BO = "BBDD Notas Crédito"
HOJA_SIEBEL = "output(1)"
