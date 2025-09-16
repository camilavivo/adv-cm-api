FROM python:3.11-slim
WORKDIR /app

# 1) Dependências
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 2) Código da API
COPY app.py cm_filler.py ./

# 3) Templates oficiais (pasta completa)
#    ➜ no seu repositório, crie a pasta "templates" na raiz e coloque dentro
#       o arquivo: Anexo 01 POP-NO-GQ-165_Rev13.docx
COPY templates/ /app/templates/

# 4) Pasta de downloads (para servir os DOCX gerados)
RUN mkdir -p /app/downloads

# 5) Variáveis de ambiente (apontando para o Rev.13)
ENV CM_TEMPLATE_PATH="/app/templates/Anexo 01 POP-NO-GQ-165_Rev13.docx"
ENV CM_DOWNLOAD_DIR="/app/downloads"
ENV CM_API_KEY=""

# 6) Expor porta (informativo)
EXPOSE 8000

# 7) Start — usar a porta do Render, com fallback 8000 para rodar local
CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
