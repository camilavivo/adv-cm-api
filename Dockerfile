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

# Garante que a pasta de downloads exista e com permissões adequadas
RUN mkdir -p /app/downloads && chmod -R 777 /app/downloads

# Variáveis de ambiente
ENV CM_TEMPLATE_PATH="/app/templates/Anexo 01 POP-NO-GQ-165_Rev13.docx"
ENV CM_DOWNLOAD_DIR="/app/downloads"
ENV CM_API_KEY=""
ENV PYTHONUNBUFFERED=1

# Expõe a porta 8000 (Render usa automaticamente a variável PORT)
EXPOSE 8000

# Comando de inicialização do servidor FastAPI (Uvicorn)
# Se o seu app principal estiver no arquivo app.py e a variável se chama "app"
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
