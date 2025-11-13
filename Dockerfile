# Imagem base leve e compatível
FROM python:3.11-slim

# Define diretório de trabalho dentro do container
WORKDIR /app

# Copia e instala dependências primeiro (melhor uso de cache)
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

# Copia os arquivos principais da aplicação
COPY app.py cm_filler.py ./ 
COPY templates/ /app/templates/

# Pasta de downloads dentro do diretório do projeto (Render permite escrita aqui)
RUN mkdir -p /opt/render/project/src/downloads && chmod -R 777 /opt/render/project/src/downloads
ENV CM_DOWNLOAD_DIR="/opt/render/project/src/downloads"

# Variáveis de ambiente
ENV CM_TEMPLATE_PATH="/app/templates/Anexo 01 POP-NO-GQ-165_Rev13.docx"
ENV CM_DOWNLOAD_DIR="/data/downloads"
ENV CM_API_KEY=""
ENV PYTHONUNBUFFERED=1

# Expõe a porta 8000 (Render usa automaticamente a variável PORT)
EXPOSE 8000

# Comando de inicialização do servidor FastAPI (Uvicorn)
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
