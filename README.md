# Azure OCR → DOCX (FastAPI + Docker)

End-to-end pipeline to convert **PDF → OCR (Azure Document Intelligence)** → **Markdown** → **DOCX**, with figure crops.  
Includes a minimal web UI for uploading PDFs and downloading the generated `.docx`.

---

## ✨ Features

- Uses **Azure Document Intelligence – prebuilt-layout** (async LRO).
- Extracts **text + figure crops** (server crops if available; falls back to local crops).
- **Pandoc** path (optional) renders LaTeX math (`$...$`, `$$...$$`) & pipe tables.
- Simple **FastAPI** service:
  - `GET /health` – health check
  - `POST /ocr-file` – multipart upload (`file`), returns `.docx`
- Minimal **web UI** (`web/index.html`) – upload PDF → download DOCX.

---

## 🚀 Quick Start (TL;DR)

```bash
# 1) Clone
git clone https://github.com/<you>/azure-ocr-docx.git
cd azure-ocr-docx

# 2) Configure Azure env (copy template and fill)
cp .env.example .env     # On Windows, copy the file manually
# edit .env with your DI endpoint & key

# 3) Build & run the API
docker compose up -d --build

# 4) Health check (should return "ok")
curl http://localhost:8000/health

# 5) Convert a PDF (CLI)
curl -X POST "http://localhost:8000/ocr-file" \
  -F "file=@samples/Patna_NewDelhi.pdf" \
  -F "title=My DOCX" \
  -F "prefer_pandoc=true" \
  --output output.docx
```

Windows/PowerShell: Use curl.exe and forward slashes in paths:

curl.exe -X POST "http://localhost:8000/ocr-file" `
  -F "file=@samples/Patna_NewDelhi.pdf" `
  -F "title=My DOCX" `
  -F "prefer_pandoc=true" `
  --output output.docx

  📦 Requirements

Docker Desktop

An Azure Document Intelligence resource:

Endpoint like: https://<your-resource>.cognitiveservices.azure.com

Key: xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

🌐 API
GET /health

Simple health probe.

POST /ocr-file — multipart/form-data

Fields

file (required): the PDF (form field must be named file)

title (optional, default My DOCX)

prefer_pandoc (optional: true/false; default false)

Response: application/vnd.openxmlformats-officedocument.wordprocessingml.document (downloadable .docx)

🖥️ Web UI

Serve the web folder (don’t open the file directly):

cd .

python -m http.server 5173
# open http://localhost:5173




