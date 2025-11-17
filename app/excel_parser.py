from pathlib import Path
from typing import Dict, List

from openpyxl import load_workbook
import unicodedata
from datetime import date, datetime


def _normalizar(texto: str) -> str:
    """
    Convierte a mayúsculas, quita tildes y espacios extra.
    Ej: 'Calificación del registro' -> 'CALIFICACION DEL REGISTRO'
    """
    if texto is None:
        return ""
    # Quitar tildes
    nfkd = unicodedata.normalize("NFD", str(texto))
    sin_tildes = "".join(c for c in nfkd if unicodedata.category(c) != "Mn")
    # Mayúsculas y sin espacios sobrantes
    return " ".join(sin_tildes.upper().split())


def parse_auditoria_excel(path: Path) -> Dict[str, List[str]]:
    """
    Lee un Excel con el formato de auditoría y devuelve los datos
    necesarios para llenar una tabla del informe.

    Devuelve un diccionario con:
      - hc: lista de códigos de historia clínica
      - fechas: lista de fechas (string)
      - porcentajes: lista de % cumplimiento
      - calificaciones: lista de calificación del registro
    """
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    fila_hc = fila_fecha = fila_porc = fila_calif = None

    # Buscamos en la columna B (columna 2) los textos clave
    for row in range(1, ws.max_row + 1):
        raw_value = ws.cell(row=row, column=2).value
        txt = _normalizar(raw_value)

        if not txt:
            continue

        # Hacemos las condiciones más flexibles
        if "CODIFICACION" in txt and "HISTORIA" in txt:
            fila_hc = row
        elif "FECHA" in txt and "ATENCION" in txt:
            fila_fecha = row
        elif "% CUMPL" in txt or "PORCENTAJE DE CUMPLIMIENTO" in txt:
            fila_porc = row
        elif "CALIFICACION" in txt and "REGISTRO" in txt:
            fila_calif = row

    if None in (fila_hc, fila_fecha, fila_porc, fila_calif):
        raise ValueError(
            f"No se encontraron todas las filas (HC, fecha, %, calificación) en {path.name}"
        )

    # Tomamos los datos desde la columna 3 hacia la derecha
    col = 3
    hc_list: List[str] = []
    fecha_list: List[str] = []
    porc_list: List[str] = []
    calif_list: List[str] = []

    while col <= ws.max_column:
        hc_val = ws.cell(row=fila_hc, column=col).value
        fecha_val = ws.cell(row=fila_fecha, column=col).value
        porc_val = ws.cell(row=fila_porc, column=col).value
        calif_val = ws.cell(row=fila_calif, column=col).value

        # Si ya no hay HC, asumimos que se acabaron las columnas útiles
        if hc_val in (None, ""):
            break

        hc_list.append(str(hc_val).strip())

        # Fecha → string amigable
        if isinstance(fecha_val, (date, datetime)):
            fecha_list.append(fecha_val.strftime("%d-%b-%y"))
        elif fecha_val is None:
            fecha_list.append("")
        else:
            fecha_list.append(str(fecha_val).strip())

        porc_list.append("" if porc_val is None else str(porc_val).strip())
        calif_list.append("" if calif_val is None else str(calif_val).strip())

        col += 1

    return {
        "hc": hc_list,
        "fechas": fecha_list,
        "porcentajes": porc_list,
        "calificaciones": calif_list,
    }
