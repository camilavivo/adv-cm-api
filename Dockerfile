FROM python:3.11-slim
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py cm_filler.py ./
COPY templates/ /app/templates/

RUN mkdir -p /app/downloads

ENV CM_TEMPLATE_PATH="/app/templates/Anexo 01 POP-NO-GQ-165_Rev13.doc"
ENV CM_DOWNLOAD_DIR="/app/downloads"
ENV CM_API_KEY=""

EXPOSE 8000
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
