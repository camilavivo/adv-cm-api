FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py cm_filler.py ./
COPY ["TESTE_FORMULARIO_CM.docx", "/app/TESTE_FORMULARIO_CM.docx"]
ENV CM_TEMPLATE_PATH="/app/TESTE_FORMULARIO_CM.docx"
ENV CM_API_KEY=""
EXPOSE 8000
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8000"]
