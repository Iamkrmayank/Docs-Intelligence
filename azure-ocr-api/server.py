from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import StreamingResponse, JSONResponse
import tempfile, subprocess, os, base64, pathlib, io
from fastapi.middleware.cors import CORSMiddleware



app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://127.0.0.1:5500",
        "http://localhost:5500",
        "http://127.0.0.1:5173",
        "http://localhost:5173",
        "http://localhost"        # optional
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def run_processor(in_pdf_path: str, title: str, prefer_pandoc: bool):
    workdir = tempfile.mkdtemp(prefix="ocr_")
    out_base = os.path.join(workdir, "my_output")
    cmd = [
        "python", "/app/test--azure.py", in_pdf_path,
        "--title", title, "--out", out_base,
        "--image-max-width", "6.5", "--embed-page-snapshots-when-no-figures"
    ]
    if prefer_pandoc:
        cmd.append("--prefer-pandoc")

    proc = subprocess.run(cmd, capture_output=True, text=True)
    print(proc.stdout); print(proc.stderr)
    if proc.returncode != 0:
        raise HTTPException(500, f"Script failed: {proc.stderr[:2000]}")

    docx_path = f"{out_base}_ocr.docx"
    if not os.path.exists(docx_path):
        raise HTTPException(500, "DOCX not found after processing")
    return docx_path

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/ocr")
async def ocr_json(pdf_b64: str, title: str = "My DOCX", prefer_pandoc: bool = False):
    # write PDF temp
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(base64.b64decode(pdf_b64))
        in_pdf = f.name
    docx_path = run_processor(in_pdf, title, prefer_pandoc)
    return StreamingResponse(open(docx_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="output.docx"'}
    )

@app.post("/ocr-file")
async def ocr_file(
    file: UploadFile = File(...),
    title: str = Form("My DOCX"),
    prefer_pandoc: bool = Form(False),
):
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        content = await file.read()
        f.write(content)
        in_pdf = f.name
    docx_path = run_processor(in_pdf, title, prefer_pandoc)
    return StreamingResponse(open(docx_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="output.docx"'}
    )
