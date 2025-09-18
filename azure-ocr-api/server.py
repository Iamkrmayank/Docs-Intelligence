from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Header
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
import tempfile, subprocess, os, base64

# --- secrets from Azure Container Apps (Env Vars bound to Secrets) ---
API_KEY = os.getenv("API_KEY", "")
AZURE_DI_KEY = os.getenv("AZURE_DI_KEY", "")              # if your test--azure.py uses it
AZURE_DI_ENDPOINT = os.getenv("AZURE_DI_ENDPOINT", "")    # same ^
ALLOWED = os.getenv("ALLOWED_ORIGINS", "")
# allow list can be comma/space separated
ALLOW_ORIGINS = [o.strip() for o in ALLOWED.replace(";", ",").split(",") if o.strip()]
if not ALLOW_ORIGINS:  # safe local defaults
    ALLOW_ORIGINS = ["http://localhost:5173", "http://127.0.0.1:5173",
                     "http://localhost:5500", "http://127.0.0.1:5500"]

app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOW_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def _require_key(x_api_key: str | None):
    if not API_KEY or not x_api_key or x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="Unauthorized")

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

    # pass DI env to child if that script expects them
    env = os.environ.copy()
    env["AZURE_DI_KEY"] = AZURE_DI_KEY
    env["AZURE_DI_ENDPOINT"] = AZURE_DI_ENDPOINT

    proc = subprocess.run(cmd, capture_output=True, text=True, env=env)
    print(proc.stdout); print(proc.stderr)
    if proc.returncode != 0:
        raise HTTPException(500, f"Script failed: {proc.stderr[:2000]}")

    docx_path = f"{out_base}_ocr.docx"
    if not os.path.exists(docx_path):
        raise HTTPException(500, "DOCX not found after processing")
    return docx_path

@app.get("/healthz")
def healthz():
    return {"ok": True}

@app.post("/ocr")
async def ocr_json(
    pdf_b64: str,
    title: str = "My DOCX",
    prefer_pandoc: bool = False,
    x_api_key: str | None = Header(None)
):
    _require_key(x_api_key)
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        f.write(base64.b64decode(pdf_b64))
        in_pdf = f.name
    docx_path = run_processor(in_pdf, title, prefer_pandoc)
    return StreamingResponse(
        open(docx_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="output.docx"'},
    )

@app.post("/ocr-file")
async def ocr_file(
    file: UploadFile = File(...),
    title: str = Form("My DOCX"),
    prefer_pandoc: bool = Form(False),
    x_api_key: str | None = Header(None)
):
    _require_key(x_api_key)
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f:
        content = await file.read()
        f.write(content)
        in_pdf = f.name
    docx_path = run_processor(in_pdf, title, prefer_pandoc)
    return StreamingResponse(
        open(docx_path, "rb"),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": 'attachment; filename="output.docx"'},
    )
