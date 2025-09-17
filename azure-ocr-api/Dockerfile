FROM python:3.11-slim

# system deps (pandoc optional; fitz often needs libgl1)
RUN apt-get update && apt-get install -y --no-install-recommends \
    pandoc libgl1 \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# copy server and your OCR script
COPY server.py test--azure.py /app/

EXPOSE 8000
CMD ["uvicorn", "server:app", "--host", "0.0.0.0", "--port", "8000"]
