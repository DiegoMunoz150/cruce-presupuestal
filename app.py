from fastapi import FastAPI, UploadFile, File, Request
from fastapi.responses import FileResponse
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
import shutil
import os

from motor_excel import generar_cruce

app = FastAPI()

templates = Jinja2Templates(directory="templates")

app.mount(
    "/static",
    StaticFiles(directory="static"),
    name="static"
)

@app.get("/")
async def inicio(request: Request):
    return templates.TemplateResponse(
        request=request,
        name="index.html"
    )


@app.post("/generar")
async def generar(
    ingresos: UploadFile = File(...),
    egresos: UploadFile = File(...)
):

    ruta_ingresos = os.path.join(
        "uploads",
        "ingresos.xlsx"
    )

    ruta_egresos = os.path.join(
        "uploads",
        "egresos.xlsx"
    )

    ruta_salida = os.path.join(
        "outputs",
        "CRUCE_FINAL.xlsx"
    )

    with open(ruta_ingresos, "wb") as buffer:
        shutil.copyfileobj(
            ingresos.file,
            buffer
        )

    with open(ruta_egresos, "wb") as buffer:
        shutil.copyfileobj(
            egresos.file,
            buffer
        )

    generar_cruce(
        ruta_ingresos,
        ruta_egresos,
        ruta_salida
    )

    return FileResponse(
        ruta_salida,
        filename="CRUCE_FINAL.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

