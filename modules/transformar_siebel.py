
def transformar_descripcion(descripcion):

    # Validar que exista una descripción
    if descripcion is None or str(descripcion).strip() == "":
        return None

    descripcion = str(descripcion).strip()

    # Separar los segmentos por /
    partes = descripcion.split("/")

    # Se esperan al menos 3 segmentos:
    # Código completo / Área-Tipo / Aprobador
    if len(partes) < 3:
        return None

    codigo_completo = partes[0].strip()
    area_tipo = partes[1].strip()
    aprobador = partes[2].strip()

    # Validar código completo, ejemplo F2-1
    if "-" not in codigo_completo:
        return None

    cod_nc = codigo_completo.split("-")[0].strip()

    # Separar área y tipo, ejemplo RETER-N
    if "-" not in area_tipo:
        return None

    area, tipo_codigo = area_tipo.rsplit("-", 1)

    area = area.strip()
    tipo_codigo = tipo_codigo.strip().upper()

    # Determinar tipo
    if tipo_codigo == "N":
        tipo = "NEGOCIO"

    elif tipo_codigo == "C":
        tipo = "CLIENTE"

    else:
        return None

    # Devolver información transformada
    return {
        "codigo_completo": codigo_completo,
        "cod_nc": cod_nc,
        "area": area,
        "tipo": tipo,
        "aprobador": aprobador,
        "descripcion_original": descripcion
    }
