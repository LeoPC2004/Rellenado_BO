
import csv
from openpyxl import load_workbook
from datetime import datetime

from config import HOJA_BO
from modules.transformar_siebel import transformar_descripcion

CONTROL_RPA = "Control_RPA"
ENCABEZADOS_CONTROL_RPA = [
    "Num_Actividad",
    "Descripcion_Siebel_Original",
    "Responsable_Siebel_Original",
    "Descripcion_Act",
    "Motivo Cliente",
    "Detalle Motivo",
    "\u00c1rea_Aprueba",
    "Aprobador",
    "Responsable",
    "Estado_Procesamiento",
    "Observacion"
]


# ============================================================
# MAPAS DE MOTIVO
# ============================================================

MAPA_MOTIVO = {
    "A1": "ANULACIÓN DE LA OPERACIÓN",
    "A2": "ANULACIÓN POR ERROR EN EL RUC",
    "D1": "DESCUENTO GLOBAL",
    "D2": "DESCUENTO POR ÍTEM",
    "E1": "DEVOLUCIÓN POR ÍTEM",
    "E2": "DEVOLUCIÓN TOTAL",
    "F2": "DISMINUCIÓN EN EL VALOR",
}


MAPA_DETALLE = {
    "A1-1": "Baja de servicio",
    "A1-2": "Baja de Servicio No Tramitada",
    "A1-3": "Cliente no reconoce contrato o modalidad",
    "A1-4": "Cambio o corrección de razón social",
    "A1-5": "Duplicidad de facturación",
    "A1-6": "Trabajos no realizados o con daños en el local",
    "A1-7": "Servicio no brindado en su totalidad",
    "A1-9": "Consolidación de comprobante",
    "A1-10": "Cambio de dirección fiscal",
    "A1-11": "Cambio de fecha de emisión",
    "A1-12": "Documento reprocesado por tratativa especial con cliente",

    "A2-1": "Cambio / Error de N° RUC",

    "D1-1": "Cliente Solicita descuentos ofrecidos al momento de firmar el contrato",
    "D1-2": "Descuento Global por conciliación",

    "D2-1": "Acuerdo de pago parcial",
    "D2-2": "Campañas Especiales",

    "E1-1": "Equipo/dispositivo devuelto",
    "E2-1": "Equipo/dispositivo devuelto",

    "F2-1": "Corrección de precios del comprobante",
    "F2-2": "Equipo/dispositivo no instalado",
    "F2-3": "La factura es mayor a la OC o Cotización",
    "F2-4": "Servicio brindado parcialmente",
}


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def esta_vacio(valor):

    return valor is None or str(valor).strip() == ""


def es_pendiente(valor):

    if esta_vacio(valor):
        return True

    return str(valor).strip().upper() == "POR DEFINIR"


def buscar_columna(ws, encabezado):

    for columna in range(1, ws.max_column + 1):

        valor = ws.cell(1, columna).value

        if valor is not None:
            if str(valor).strip() == encabezado:
                return columna

    raise ValueError(
        f"No se encontró la columna '{encabezado}' "
        f"en la hoja {ws.title}"
    )


# ============================================================
# PROCESO PRINCIPAL
# ============================================================

def rellenar_bo(
    ruta_bo,
    ruta_siebel,
    ruta_salida
):

    print("\n" + "=" * 60)
    print("PROCESO DE RELLENADO DEL BO")
    print("=" * 60)

    # --------------------------------------------------------
    # 1. ABRIR BO
    # --------------------------------------------------------

    wb_bo = load_workbook(ruta_bo)
    ws_bo = wb_bo[HOJA_BO]

    # --------------------------------------------------------
    # 2. DETECTAR ACTIVIDADES PENDIENTES
    # --------------------------------------------------------

    actividades_pendientes = []

    for fila in range(2, ws_bo.max_row + 1):

        actividad = ws_bo.cell(fila, 11).value  # K
        descripcion_act = ws_bo.cell(fila, 27).value  # AA

        if es_pendiente(descripcion_act):

            if not esta_vacio(actividad):

                actividad = str(actividad).strip()

                actividades_pendientes.append({
                    "fila": fila,
                    "actividad": actividad
                })

    print(
        f"\nActividades pendientes detectadas: "
        f"{len(actividades_pendientes)}"
    )

    # --------------------------------------------------------
    # 3. REGISTRAR ACTIVIDADES EN MEMORIA Y GENERAR OR
    # --------------------------------------------------------

    lista_actividades = [
        elemento["actividad"]
        for elemento in actividades_pendientes
    ]

    consulta_or = " or ".join(lista_actividades)

    print("\nCONSULTA QUE SE PEGARÍA EN SIEBEL:")
    print("-" * 60)
    print(consulta_or)
    print("-" * 60)

    # --------------------------------------------------------
    # 4. SIMULAR EXTRACCIÓN: LEER EXCEL DE SIEBEL
    # --------------------------------------------------------

    with open(ruta_siebel, "r", encoding="utf-16", newline="") as archivo_siebel:
        lector_siebel = csv.DictReader(
            archivo_siebel,
            delimiter="\t"
        )

        columnas_requeridas = [
            "N\u00ba de actividad",
            "Descripci\u00f3n",
            "Responsable"
        ]

        encabezados = lector_siebel.fieldnames or []
        encabezados = [encabezado.strip() for encabezado in encabezados]
        lector_siebel.fieldnames = encabezados
        columnas_faltantes = [
            columna
            for columna in columnas_requeridas
            if columna not in encabezados
        ]

        if columnas_faltantes:
            raise ValueError(
                "Faltan columnas en el export de Siebel: "
                + ", ".join(columnas_faltantes)
            )

        datos_siebel = {}
        duplicados_siebel = []

        for registro in lector_siebel:

            actividad = registro["N\u00ba de actividad"]

            if esta_vacio(actividad):
                continue

            actividad = str(actividad).strip()
            descripcion = registro["Descripci\u00f3n"]
            responsable = registro["Responsable"]

            if actividad in datos_siebel:

                duplicados_siebel.append(actividad)
                continue

            datos_siebel[actividad] = {
                "descripcion": descripcion,
                "responsable": responsable
            }

    print(
        f"\nRegistros leídos del export de Siebel: "
        f"{len(datos_siebel)}"
    )

    if duplicados_siebel:

        print("\nADVERTENCIA: Actividades duplicadas en Siebel:")

        for actividad in duplicados_siebel:
            print(f" - {actividad}")

    # --------------------------------------------------------
    # 6. DEFINIR COLUMNAS DEL BO
    # --------------------------------------------------------

    # K  = Num_Actividad
    # AA = Descripcion_Act
    # AB = Motivo Cliente
    # AC = Detalle Motivo
    # AD = Área_Aprueba
    # AE = Aprobador
    # AK = Responsable

    COL_ACTIVIDAD = 11
    COL_DESCRIPCION = 27
    COL_MOTIVO = 28
    COL_DETALLE = 29
    COL_AREA = 30
    COL_APROBADOR = 31
    COL_RESPONSABLE = 37

    # --------------------------------------------------------
    # 7. CRUZAR, RELLENAR Y REGISTRAR AUDITORÍA
    # --------------------------------------------------------

    cambios = []

    actividades_no_encontradas = []
    descripciones_invalidas = []

    filas_procesadas = 0
    registros_control = []

    columnas_control = {
        COL_DESCRIPCION: "Descripcion_Act",
        COL_MOTIVO: "Motivo Cliente",
        COL_DETALLE: "Detalle Motivo",
        COL_AREA: "\u00c1rea_Aprueba",
        COL_APROBADOR: "Aprobador",
        COL_RESPONSABLE: "Responsable"
    }

    for elemento in actividades_pendientes:

        fila_bo = elemento["fila"]
        actividad = elemento["actividad"]

        print(f"\nProcesando: {actividad}")

        registro_control = {
            encabezado: None
            for encabezado in ENCABEZADOS_CONTROL_RPA
        }
        registro_control["Num_Actividad"] = actividad
        observaciones = []
        datos_incompletos = []

        # Buscar en el export de Siebel
        registro_siebel = datos_siebel.get(actividad)

        if registro_siebel is None:

            print("  No se encontr\u00f3 en el export de Siebel.")
            actividades_no_encontradas.append(actividad)
            registro_control["Estado_Procesamiento"] = "ERROR"
            registro_control["Observacion"] = (
                "Actividad no encontrada en el export de Siebel."
            )
            registros_control.append(registro_control)
            continue

        descripcion = registro_siebel["descripcion"]
        responsable_siebel = registro_siebel["responsable"]
        registro_control["Descripcion_Siebel_Original"] = descripcion
        registro_control["Responsable_Siebel_Original"] = responsable_siebel

        if esta_vacio(descripcion):
            observaciones.append("Descripción de Siebel vacía.")
            datos_incompletos.append(True)

        if esta_vacio(responsable_siebel):
            observaciones.append("Responsable original de Siebel vacío.")
            datos_incompletos.append(True)

        # Transformar descripción con la lógica existente
        datos_transformados = transformar_descripcion(
            descripcion
        )

        if datos_transformados is None:

            print("  Descripción inválida o no reconocida.")
            descripciones_invalidas.append(actividad)
            registro_control["Estado_Procesamiento"] = "INCOMPLETO"
            observaciones.append(
                "No se pudo transformar la descripción de Siebel."
            )
            registro_control["Observacion"] = " ".join(observaciones)
            registros_control.append(registro_control)
            continue

        filas_procesadas += 1

        # ----------------------------------------------------
        # PREPARAR VALORES
        # ----------------------------------------------------

        valores_nuevos = {

            COL_DESCRIPCION:
                datos_transformados["descripcion_original"],

            COL_MOTIVO:
                MAPA_MOTIVO.get(
                    datos_transformados["cod_nc"]
                ),

            COL_DETALLE:
                MAPA_DETALLE.get(
                    datos_transformados["codigo_completo"]
                ),

            COL_AREA:
                datos_transformados["area"],

            COL_APROBADOR:
                datos_transformados["aprobador"],

            COL_RESPONSABLE:
                (
                    datos_transformados["tipo"]
                )
        }

        for columna, valor_nuevo in valores_nuevos.items():

            encabezado_control = columnas_control[columna]

            if valor_nuevo is None or esta_vacio(valor_nuevo):
                observaciones.append(
                    f"No se obtuvo un valor para {encabezado_control}."
                )
                datos_incompletos.append(True)
            else:
                # Es el resultado que la lógica intenta asignar al BO.
                registro_control[encabezado_control] = valor_nuevo

            if valor_nuevo is None:
                continue

            celda = ws_bo.cell(
                fila_bo,
                columna
            )

            if es_pendiente(celda.value):

                valor_anterior = celda.value

                celda.value = valor_nuevo

                cambios.append({
                    "fila": fila_bo,
                    "actividad": actividad,
                    "columna": ws_bo.cell(
                        1,
                        columna
                    ).value,
                    "valor_anterior": valor_anterior,
                    "valor_nuevo": valor_nuevo
                })

                print(
                    f"  Modificado: "
                    f"{ws_bo.cell(1, columna).value} "
                    f"-> {valor_nuevo}"
                )

            else:

                observaciones.append(
                    f"{encabezado_control}: BO conservó el valor existente; "
                    "no se sobrescribió."
                )
                print(
                    f"  Sin cambios: "
                    f"{ws_bo.cell(1, columna).value} "
                    f"(ya tiene valor)"
                )

        registro_control["Estado_Procesamiento"] = (
            "INCOMPLETO" if datos_incompletos else "COMPLETADO"
        )
        registro_control["Observacion"] = (
            " ".join(observaciones) if observaciones else "Procesamiento correcto."
        )
        registros_control.append(registro_control)

    # Reemplazar solo la hoja de auditoría para no conservar filas anteriores.
    if CONTROL_RPA in wb_bo.sheetnames:
        del wb_bo[CONTROL_RPA]

    ws_control = wb_bo.create_sheet(CONTROL_RPA)
    ws_control.append(ENCABEZADOS_CONTROL_RPA)

    for registro_control in registros_control:
        ws_control.append([
            registro_control[encabezado]
            for encabezado in ENCABEZADOS_CONTROL_RPA
        ])

    # --------------------------------------------------------
    # 8. GUARDAR BO RESULTANTE Y AUDITORÍA
    # --------------------------------------------------------
    wb_bo.save(ruta_salida)
    wb_bo.close()

    # --------------------------------------------------------
    # 9. RESUMEN
    # --------------------------------------------------------

    print("\n" + "=" * 60)
    print("RESUMEN DEL PROCESO")
    print("=" * 60)

    print(
        f"Actividades pendientes detectadas: "
        f"{len(actividades_pendientes)}"
    )

    print(
        f"Actividades encontradas en Siebel: "
        f"{filas_procesadas}"
    )

    print(
        f"Modificaciones realizadas: "
        f"{len(cambios)}"
    )

    print(
        f"Actividades no encontradas: "
        f"{len(actividades_no_encontradas)}"
    )

    print(
        f"Descripciones inválidas: "
        f"{len(descripciones_invalidas)}"
    )

    print(f"\nArchivo generado: {ruta_salida}")

    if actividades_no_encontradas:

        print("\nACTIVIDADES NO ENCONTRADAS:")

        for actividad in actividades_no_encontradas:
            print(f" - {actividad}")

    if descripciones_invalidas:

        print("\nDESCRIPCIONES INVÁLIDAS:")

        for actividad in descripciones_invalidas:
            print(f" - {actividad}")

    return cambios
