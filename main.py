
import os
import sys
import time
from datetime import datetime

from openpyxl import Workbook, load_workbook

import config
from modules.coordenadas import obtener_export_siebel
from modules.rellenar_bo import rellenar_bo


# ============================================================
# OBTENER ACTIVIDADES PENDIENTES DEL BO
# ============================================================

def obtener_actividades_pendientes(ruta_bo):

    print("\nRevisando actividades pendientes del BO...")

    wb = load_workbook(
        ruta_bo,
        read_only=True,
        data_only=True
    )

    ws = wb[config.HOJA_BO]

    actividades = []

    # K = Num_Actividad
    # AA = Descripcion_Act

    for fila in range(2, ws.max_row + 1):

        num_actividad = ws.cell(fila, 11).value
        descripcion = ws.cell(fila, 27).value

        pendiente = (
            descripcion is None
            or str(descripcion).strip() == ""
            or str(descripcion).strip().upper() == "POR DEFINIR"
        )

        if pendiente and num_actividad is not None:

            actividad = str(num_actividad).strip()

            if actividad:
                actividades.append(actividad)

    wb.close()

    # Eliminar IDs repetidos manteniendo el orden
    actividades = list(dict.fromkeys(actividades))

    print(f"Actividades pendientes: {len(actividades)}")

    return actividades


# ============================================================
# CONSTRUIR CONSULTA OR
# ============================================================

def construir_consulta_or(actividades):

    codigo_or = " OR ".join(actividades)

    return codigo_or


# ============================================================
# REGISTRO DE MODIFICACIONES
# ============================================================

def guardar_registro(cambios, ruta_registro):

    if not cambios:

        print("\nNo hay modificaciones que registrar.")
        return

    fecha_ejecucion = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )

    # Si existe el registro, abrirlo
    if os.path.exists(ruta_registro):

        wb = load_workbook(ruta_registro)
        ws = wb.active

    else:

        # Crear registro nuevo
        wb = Workbook()
        ws = wb.active
        ws.title = "Modificaciones"

        ws.append([
            "Fecha ejecución",
            "Num_Actividad",
            "Columna modificada",
            "Valor anterior",
            "Valor nuevo"
        ])

    # Agregar modificaciones
    for cambio in cambios:

        ws.append([
            fecha_ejecucion,
            cambio["actividad"],
            cambio["columna"],
            cambio["valor_anterior"],
            cambio["valor_nuevo"]
        ])

    wb.save(ruta_registro)
    wb.close()

    print(f"\nRegistro actualizado: {ruta_registro}")


# ============================================================
# VALIDAR EXPORT DE SIEBEL
# ============================================================

def validar_export_siebel(ruta_siebel):

    print("\nValidando export de Siebel...")

    if not os.path.exists(ruta_siebel):
        raise FileNotFoundError(
            f"No se generó el export de Siebel: {ruta_siebel}"
        )

    if os.path.getsize(ruta_siebel) == 0:
        raise ValueError(
            f"El export de Siebel está vacío: {ruta_siebel}"
        )

    print("Export de Siebel validado correctamente.")
# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 60)
    print("RPA RELLENADO BO")
    print("EXTRACCIÓN AUTOMÁTICA DESDE SIEBEL")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. VALIDAR BO ORIGINAL
    # --------------------------------------------------------

    if not os.path.exists(config.RUTA_BO):

        print("\nERROR: No se encontró el archivo BO.")
        print(config.RUTA_BO)
        return

    print("\nBO encontrado:")
    print(config.RUTA_BO)

    # --------------------------------------------------------
    # 2. IDENTIFICAR ACTIVIDADES PENDIENTES
    # --------------------------------------------------------

    actividades = obtener_actividades_pendientes(
        config.RUTA_BO
    )

    if not actividades:

        print("\nNo se encontraron actividades pendientes.")
        print("No es necesario ejecutar Siebel.")

        return

    print(
        f"\nSe encontraron {len(actividades)} "
        "actividades pendientes."
    )

    # --------------------------------------------------------
    # 3. CONSTRUIR CONSULTA OR
    # --------------------------------------------------------

    codigo_or = construir_consulta_or(actividades)

    print("\nConsulta preparada para Siebel.")

    # --------------------------------------------------------
    # 4. PREPARAR RUTA ÚNICA PARA LA EXPORTACIÓN
    # --------------------------------------------------------

    ruta_siebel = config.generar_ruta_export_siebel()
    print(f"\nEl export de esta ejecución se guardará en: {ruta_siebel}")

    # --------------------------------------------------------
    # 5. EXTRAER INFORMACIÓN DESDE SIEBEL
    # --------------------------------------------------------

    print("\n" + "=" * 60)
    print("INICIANDO AUTOMATIZACIÓN DE SIEBEL")
    print("=" * 60)

    obtener_export_siebel(codigo_or, ruta_siebel)

    # --------------------------------------------------------
    # 6. VALIDAR EXPORT NUEVO
    # --------------------------------------------------------

    validar_export_siebel(ruta_siebel)

    # --------------------------------------------------------
    # 7. COMPARAR Y RELLENAR BO
    # --------------------------------------------------------

    print("\n" + "=" * 60)
    print("INICIANDO RELLENADO DEL BO")
    print("=" * 60)

    cambios = rellenar_bo(
        ruta_bo=config.RUTA_BO,
        ruta_siebel=ruta_siebel,
        ruta_salida=config.RUTA_BO_RELLENADO
    )

    # --------------------------------------------------------
    # 8. REGISTRAR MODIFICACIONES
    # --------------------------------------------------------

    guardar_registro(
        cambios,
        config.RUTA_REGISTRO
    )

    # --------------------------------------------------------
    # 9. FINALIZACIÓN
    # --------------------------------------------------------

    print("\n" + "=" * 60)
    print("PROCESO FINALIZADO")
    print("=" * 60)

    print("\nArchivos generados:")

    print("\nExport de Siebel:")
    print(ruta_siebel)

    print("\nBO rellenado:")
    print(config.RUTA_BO_RELLENADO)

    print("\nRegistro de modificaciones:")
    print(config.RUTA_REGISTRO)


# ============================================================
# EJECUCIÓN
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except KeyboardInterrupt:

        print("\nProceso interrumpido por el usuario.")

    except Exception as e:

        print("\n" + "=" * 60)
        print("ERROR DURANTE EL PROCESO")
        print("=" * 60)

        print(e)

        sys.exit(1)
