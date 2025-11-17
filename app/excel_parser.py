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


def _encontrar_hoja_correcta(wb):
    """
    Busca en todas las hojas del workbook la que contenga el texto
    'FORMATO DE EVALUACIÓN DE LA CALIDAD DE REGISTO EN CONSULTA EXTERNA'
    (con o sin tildes)
    """
    texto_buscar = "FORMATO DE EVALUACION DE LA CALIDAD DE REGISTO EN CONSULTA EXTERNA"
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        
        # Buscar en las primeras filas (usualmente el título está arriba)
        for row in range(1, min(20, ws.max_row + 1)):  # Buscar en las primeras 20 filas
            for col in range(1, min(10, ws.max_column + 1)):  # Primeras 10 columnas
                cell_value = ws.cell(row=row, column=col).value
                if cell_value:
                    texto_normalizado = _normalizar(cell_value)
                    # Verificar si contiene las palabras clave
                    if "FORMATO" in texto_normalizado and "EVALUACION" in texto_normalizado and "CALIDAD" in texto_normalizado:
                        return ws
    
    return None


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
    
    print(f"\n=== Procesando archivo: {path.name} ===")
    print(f"Hojas disponibles: {wb.sheetnames}")
    
    # Buscar la hoja correcta
    ws = _encontrar_hoja_correcta(wb)
    
    if ws is None:
        raise ValueError(
            f"No se encontró la hoja con el formato esperado en {path.name}. "
            f"Se buscó una hoja que contenga 'FORMATO DE EVALUACIÓN DE LA CALIDAD DE REGISTO EN CONSULTA EXTERNA'"
        )
    
    print(f"✓ Hoja encontrada: {ws.title}")

    fila_hc = fila_fecha = fila_porc = fila_calif = None
    col_inicio = None

    print("\nBuscando las 4 filas clave en TODO el Excel...")
    # Buscar en todas las celdas las filas que contienen estos textos
    for row in range(1, min(200, ws.max_row + 1)):
        for col in range(1, min(20, ws.max_column + 1)):
            raw_value = ws.cell(row=row, column=col).value
            txt = _normalizar(raw_value)

            if not txt:
                continue

            # Mostrar las primeras celdas para debug
            if row <= 100 and col <= 3:
                print(f"  Fila {row}, Col {col}: {raw_value}")

            # Buscar exactamente estos textos
            if "N" in txt and "HC" in txt and "EVALUADA" in txt:
                fila_hc = row
                if col_inicio is None:
                    col_inicio = col
                print(f"  ✓✓✓ Fila HC encontrada en: fila={row}, col={col}")
            
            if "FECHA" in txt and "ATENCION" in txt and "EVALUADA" in txt:
                fila_fecha = row
                if col_inicio is None:
                    col_inicio = col
                print(f"  ✓✓✓ Fila FECHA encontrada en: fila={row}, col={col}")
            
            if "%" in txt and "CUMPLIMIENTO" in txt:
                fila_porc = row
                if col_inicio is None:
                    col_inicio = col
                print(f"  ✓✓✓ Fila PORCENTAJE encontrada en: fila={row}, col={col}")
            
            if "CALIFICACION" in txt and "REGISTRO" in txt:
                fila_calif = row
                if col_inicio is None:
                    col_inicio = col
                print(f"  ✓✓✓ Fila CALIFICACIÓN encontrada en: fila={row}, col={col}")

    if None in (fila_hc, fila_fecha, fila_porc, fila_calif):
        raise ValueError(
            f"No se encontraron todas las filas requeridas en {path.name}. "
            f"Encontradas - HC: {'✓' if fila_hc else '✗'}, "
            f"Fecha: {'✓' if fila_fecha else '✗'}, "
            f"Porcentaje: {'✓' if fila_porc else '✗'}, "
            f"Calificación: {'✓' if fila_calif else '✗'}"
        )

    print(f"\n✓ Todas las filas encontradas:")
    print(f"  HC en fila: {fila_hc}")
    print(f"  Fecha en fila: {fila_fecha}")
    print(f"  Porcentaje en fila: {fila_porc}")
    print(f"  Calificación en fila: {fila_calif}")
    print(f"  Columna de inicio (títulos): {col_inicio}")

    # Los datos empiezan en la siguiente columna después de los títulos
    col_datos = col_inicio + 1 if col_inicio else 2

    # Tomamos los datos desde la columna siguiente a los títulos
    col = col_datos
    hc_list: List[str] = []
    fecha_list: List[str] = []
    porc_list: List[str] = []
    calif_list: List[str] = []

    print(f"\nExtrayendo datos desde la columna {col_datos}...")
    while col <= ws.max_column:
        hc_val = ws.cell(row=fila_hc, column=col).value
        fecha_val = ws.cell(row=fila_fecha, column=col).value
        porc_val = ws.cell(row=fila_porc, column=col).value
        calif_val = ws.cell(row=fila_calif, column=col).value

        print(f"  Columna {col}: HC={hc_val}, Fecha={fecha_val}, %={porc_val}, Calif={calif_val}")

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

    print(f"\n✓ Datos extraídos:")
    print(f"  - {len(hc_list)} historias clínicas: {hc_list}")
    print(f"  - Fechas: {fecha_list}")
    print(f"  - Porcentajes: {porc_list}")
    print(f"  - Calificaciones: {calif_list}")

    return {
        "hc": hc_list,
        "fechas": fecha_list,
        "porcentajes": porc_list,
        "calificaciones": calif_list,
    }