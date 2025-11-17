from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, HTMLResponse
from pathlib import Path
import shutil

from .excel_parser import parse_auditoria_excel
from .report_generator import generar_reporte

app = FastAPI()

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

PLANTILLA_PATH = Path(__file__).resolve().parent / "plantilla_informe.docx"


@app.get("/", response_class=HTMLResponse)
def index():
    return """
    <html>
      <body>
        <h2>Generador de Informe HC</h2>
        <form action="/procesar" method="post" enctype="multipart/form-data">
          <p>Sube uno o varios archivos Excel:</p>
          <input name="files" type="file" multiple />
          <button type="submit">Procesar</button>
        </form>
      </body>
    </html>
    """


@app.post("/procesar")
async def procesar(files: list[UploadFile] = File(...)):
    try:
        saved_paths = []
        for f in files:
            dest = UPLOAD_DIR / f.filename
            with dest.open("wb") as buffer:
                shutil.copyfileobj(f.file, buffer)
            saved_paths.append(dest)

        # 1) Parsear cada Excel subido
        bloques_datos = []
        for path in saved_paths:
            bloques_datos.append(parse_auditoria_excel(path))

        # 2) Generar reporte a partir de la plantilla
        output_path = BASE_DIR / "reporte_generado.docx"
        generar_reporte(bloques_datos, PLANTILLA_PATH, output_path)

        # 3) Devolver el archivo generado
        return FileResponse(
            path=output_path,
            filename="reporte_generado.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    except Exception as e:
        # Esto hará que el navegador muestre el detalle
        raise HTTPException(status_code=400, detail=str(e))
