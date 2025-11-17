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

    # Índice de la tabla que vamos a llenar
    tabla_idx = 0

    for bloque in bloques:
        hc = bloque["hc"]
        fechas = bloque["fechas"]
        porcentajes = bloque["porcentajes"]
        calificaciones = bloque["calificaciones"]

        # Si hay tablas existentes en la plantilla, intentamos usarlas
        if tabla_idx < len(doc.tables):
            tabla = doc.tables[tabla_idx]
            
            # Verificar que la tabla tenga al menos 4 filas (para nuestros datos)
            while len(tabla.rows) < 4:
                tabla.add_row()
            
            # Calcular cuántas columnas necesitamos (datos + columna de títulos)
            num_cols_necesarias = len(hc) + 1
            
            # Agregar columnas si hacen falta
            while len(tabla.columns) < num_cols_necesarias:
                tabla.add_column(width=1)
            
            # Llenar la primera columna con los títulos (si no están ya)
            tabla.rows[0].cells[0].text = "N° de HC evaluada"
            tabla.rows[1].cells[0].text = "Fecha de atención evaluada"
            tabla.rows[2].cells[0].text = "% cumplimiento"
            tabla.rows[3].cells[0].text = "Calificación del registro"
            
            # Llenar las columnas de datos
            for idx, (h, f, p, c) in enumerate(
                zip(hc, fechas, porcentajes, calificaciones), start=1
            ):
                if idx >= len(tabla.columns):
                    tabla.add_column(width=1)
                
                tabla.rows[0].cells[idx].text = h
                tabla.rows[1].cells[idx].text = f
                tabla.rows[2].cells[idx].text = p
                tabla.rows[3].cells[idx].text = c
            
            tabla_idx += 1
        else:
            # Si ya no hay tablas en la plantilla, crear una nueva
            num_cols = len(hc) + 1
            num_cols = max(num_cols, 2)

            tabla = doc.add_table(rows=4, cols=num_cols)
            tabla.style = 'Table Grid'  # Aplicar un estilo con bordes

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