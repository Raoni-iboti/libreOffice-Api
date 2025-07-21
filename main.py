
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
import os
import shutil
import subprocess

app = FastAPI()
UPLOAD_PATH = "temp"
INPUT_FILE = os.path.join(UPLOAD_PATH, "original.docx")
OUTPUT_FILE = os.path.join(UPLOAD_PATH, "editado.docx")

os.makedirs(UPLOAD_PATH, exist_ok=True)

@app.post("/abrir")
async def abrir_docx(file: UploadFile = File(...)):
    with open(INPUT_FILE, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)
    shutil.copy(INPUT_FILE, OUTPUT_FILE)
    return {"status": "ok", "message": "Arquivo salvo"}

@app.post("/substituir")
async def substituir_texto(buscar: str = Form(...), por: str = Form(...)):
    cmd = [
        "libreoffice",
        "--headless",
        "--convert-to", "docx",
        "--outdir", UPLOAD_PATH,
        f"macro:///Standard.Module1.ReplaceText(\"{OUTPUT_FILE}\",\"{buscar}\",\"{por}\")"
    ]
    subprocess.run(cmd)
    return {"status": "ok", "message": "Texto substituído"}

@app.get("/baixar")
async def baixar():
    return FileResponse(OUTPUT_FILE, filename="editado.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
