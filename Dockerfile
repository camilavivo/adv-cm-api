FROM python:3.11-slim
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py cm_filler.py ./
COPY TESTE_FORMULARIO_CM.docx /app/TESTE_FORMULARIO_CM.docx

# cria a pasta de downloads (onde os DOCX ficarão para download público)
RUN mkdir -p /app/downloads

ENV CM_TEMPLATE_PATH="/app/TESTE_FORMULARIO_CM.docx"
# opcional: outra pasta
ENV CM_DOWNLOAD_DIR="/app/downloads"
ENV CM_API_KEY=""

EXPOSE 8000
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
