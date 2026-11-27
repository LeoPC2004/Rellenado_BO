
import time
import pyautogui
import pyperclip
import os



# =========================================================
# CONFIGURACIÓN DE COORDENADAS
# =========================================================

# Apertura de Siebel

HERRAMIENTAS = (455, 123)

SIEBEL = (545, 216)

COORDENADA_1 = (1062, 685)  # Campo usuario

COORDENADA_2 = (1223, 685)  # Campo contraseña

LOGIN = (1353, 684)

# Navegación

FACTURAS = (1806, 238)

CONSULTA_ACTIVIDADES_ARGENTINAS = (1743, 434)

MODULO_CONSULTA = (508, 269)

NUM_ACTIVIDAD = (190, 322)

IR = (436, 268)

# Exportación

MENU = (348, 267)

EXPORTAR = (368, 619)

SIGUIENTE = (1695, 847)

GUARDAR = (1047, 604)

CERRAR = (1165, 651)


# =========================================================
# APERTURA DE EDGE
# =========================================================

def abrir_edge():

    print("  - Abriendo Edge...", flush=True)

    print("Abriendo Microsoft Edge...")

    os.system("start microsoft-edge:")

    time.sleep(6)

# =========================================================
# INGRESO A SIEBEL
# =========================================================

def login_siebel():

    print("  - Accediendo a Siebel e iniciando sesion...", flush=True)
    print("Ingresando a Siebel...")

    pyautogui.click(*HERRAMIENTAS)
    time.sleep(3)

    pyautogui.click(*SIEBEL)
    time.sleep(5)

    pyautogui.doubleClick(*COORDENADA_1)
    time.sleep(2)

    pyautogui.doubleClick(*COORDENADA_2)
    time.sleep(2)

    pyautogui.click(*LOGIN)
    time.sleep(15)


# =========================================================
# NAVEGACIÓN AL MÓDULO DE CONSULTA
# =========================================================

def navegar_consulta():

    print("  - Navegando al modulo de consulta...", flush=True)
    print("Navegando al módulo de consulta...")

    pyautogui.click(*FACTURAS)
    time.sleep(5)

    pyautogui.click(*CONSULTA_ACTIVIDADES_ARGENTINAS)
    time.sleep(5)

    pyautogui.click(*MODULO_CONSULTA)
    time.sleep(5)


# =========================================================
# EJECUTAR CONSULTA
# =========================================================

def ejecutar_consulta(codigo_or):

    print("  - Ingresando actividades y ejecutando consulta...", flush=True)
    print("Ingresando actividades para la consulta...")

    pyautogui.click(*NUM_ACTIVIDAD)
    time.sleep(2)

    pyautogui.hotkey("ctrl", "a")
    pyperclip.copy(str(codigo_or))
    pyautogui.hotkey("ctrl", "v")

    time.sleep(2)

    pyautogui.click(*IR)
    time.sleep(15)


# =========================================================
# EXPORTAR RESULTADOS DE SIEBEL
# =========================================================

def exportar_siebel(ruta_siebel):

    print("  - Solicitando exportacion y esperando el archivo...", flush=True)
    print("Exportando resultados de Siebel...")

    pyautogui.click(*MENU)
    time.sleep(2)

    pyautogui.click(*EXPORTAR)
    time.sleep(5)

    pyautogui.click(*SIGUIENTE)
    time.sleep(5)

    pyautogui.click(*GUARDAR)
    time.sleep(5)

    # Cuadro de guardado de Windows
    pyautogui.hotkey("alt", "n")
    time.sleep(2)

    pyautogui.hotkey("ctrl", "a")
    time.sleep(1)

    pyperclip.copy(ruta_siebel)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(3)

    pyautogui.press("enter")

    # Esperar a que el CSV exista y su tamaño permanezca estable.
    limite_espera = time.monotonic() + 180
    tamano_anterior = None
    lecturas_estables = 0

    while time.monotonic() < limite_espera:
        if os.path.isfile(ruta_siebel):
            tamano_actual = os.path.getsize(ruta_siebel)

            if tamano_actual > 0 and tamano_actual == tamano_anterior:
                lecturas_estables += 1
                if lecturas_estables >= 2:
                    break
            else:
                lecturas_estables = 0

            tamano_anterior = tamano_actual

        time.sleep(1)
    else:
        raise TimeoutError(
            f"La exportación de Siebel no terminó a tiempo: {ruta_siebel}"
        )

    # Cerrar ventana de exportación
    pyautogui.click(*CERRAR)
    time.sleep(3)

    print("Exportación finalizada.")


# =========================================================
# PROCESO COMPLETO DE COORDENADAS
# =========================================================

def obtener_export_siebel(codigo_or, ruta_siebel):

    abrir_edge()

    login_siebel()

    navegar_consulta()

    ejecutar_consulta(codigo_or)

    exportar_siebel(ruta_siebel)

    print("Proceso de extracción de Siebel completado.")
