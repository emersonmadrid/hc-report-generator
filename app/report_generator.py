from pathlib import Path
from typing import List, Dict

from docx import Document


def generar_reporte(
    bloques: List[Dict[str, List[str]]],
    plantilla_path: Path,
    output_path: Path,
) -> None:
    """
    Crea un Word a partir de una plantilla y una lista de bloques de datos.
    Cada bloque corresponde a un Excel (una tabla en el informe).
    """
    doc = Document(plantilla_path)

    # Si quieres conservar la tabla de la plantilla, no la borres.
    # Aquí, en lugar de usar la tabla existente, vamos a agregar tablas nuevas debajo.
    for bloque in bloques:
        hc = bloque["hc"]
        fechas = bloque["fechas"]
        porcentajes = bloque["porcentajes"]
        calificaciones = bloque["calificaciones"]

        num_cols = len(hc) + 1  # +1 para la columna izquierda de títulos
        num_cols = max(num_cols, 2)  # mínimo 2 columnas

        tabla = doc.add_table(rows=4, cols=num_cols)

        # Primera columna: nombres fijos de las filas
        tabla.cell(0, 0).text = "N° de HC evaluada"
        tabla.cell(1, 0).text = "Fecha de atención evaluada"
        tabla.cell(2, 0).text = "% cumplimiento"
        tabla.cell(3, 0).text = "Calificación del registro"

        # Rellenar columnas de datos
        for idx, (h, f, p, c) in enumerate(
            zip(hc, fechas, porcentajes, calificaciones), start=1
        ):
            if idx >= num_cols:
                break

            tabla.cell(0, idx).text = h
            tabla.cell(1, idx).text = f
            tabla.cell(2, idx).text = p
            tabla.cell(3, idx).text = c

        # Espacio entre tablas
        doc.add_paragraph("")

    doc.save(output_path)
