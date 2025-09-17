# -*- coding: utf-8 -*-
import os
import uuid
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Header, Response, Request
from pydantic import BaseModel, Field
from starlette.staticfiles import StaticFiles

from cm_filler import preencher_docx_from_payload

# variáveis de ambiente (com defaults seguros)
API_KEY = os.getenv("CM_API_KEY")
TEMPLATE_PATH = os.getenv("CM_TEMPLATE_PATH", "/app/templates/Anexo 01 POP-NO-GQ-165_Rev13.docx")
DOWNLOAD_DIR = os.getenv("CM_DOWNLOAD_DIR", "/app/downloads")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

app = FastAPI(
    title="ADV CM Filler API",
    version="1.3.0",
    description="API que preenche o formulário CM e retorna o DOCX (binário, base64 e URL).",
)

# rota estática para baixar arquivos gerados
app.mount("/downloads", StaticFiles(directory=DOWNLOAD_DIR), name="downloads")


class CMPayload(BaseModel):
    # Cabeçalho
    numero_cm: Optional[str] = None  # >>> novo: será impresso no cabeçalho como "CM nº: <valor>"

    # Seção 1–3
    data: str
    solicitante: str
    departamento: str
    situacao_atual: str
    alteracao_proposta: str
    descricoes_itens: List[str] = Field(default_factory=list)
    numeros_correspondentes: List[str] = Field(default_factory=list)
    abrangencia: str
    titulo_mudanca: Optional[str] = None
    caracter_mudanca: str = "Temporária"
    retorno_mudanca_temp: Optional[str] = None
    mudanca_refere_se: List[str] = Field(default_factory=list)
    impactos: List[str] = Field(default_factory=list)
    classificacao: str
    justificativa_classificacao: str
    justificativa_mudanca: Optional[str] = ""  # preenche "JUSTIFICATIVA MUDANÇA"

    # Seção 4–7
    anexos_aplicaveis: List[str] = Field(default_factory=list)
    departamentos_pertinentes: List[str] = Field(default_factory=list)
    treinamento_executado: Optional[bool] = None
    plano_implementacao: List[str] = Field(default_factory=list)
    voe_criterios: Optional[str] = ""
    voe_periodo: Optional[str] = ""
    voe_resultados_esperados: Optional[str] = ""
    observacoes_finais: Optional[str] = ""     # vazio por padrão

def _check_api_key(x_api_key: Optional[str]):
    if not API_KEY:
        return  # sem API key configurada, não exige
    if not x_api_key or x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key inválida")

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post(
    "/fill",
    response_class=Response,
    responses={200: {"content": {"application/vnd.openxmlformats-officedocument.wordprocessingml.document": {}}}},
)
def fill(payload: CMPayload, x_api_key: Optional[str] = Header(default=None)):
    _check_api_key(x_api_key)
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template não encontrado: {TEMPLATE_PATH}")
    docx_bytes = preencher_docx_from_payload(TEMPLATE_PATH, payload.model_dump())
    headers = {"Content-Disposition": 'attachment; filename="FORMULARIO_CM_PRENCHIDO.docx"'}
    return Response(
        content=docx_bytes,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers,
    )

@app.post("/fill_b64")
def fill_b64(payload: CMPayload, x_api_key: Optional[str] = Header(default=None)):
    import base64
    _check_api_key(x_api_key)
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template não encontrado: {TEMPLATE_PATH}")
    docx_bytes = preencher_docx_from_payload(TEMPLATE_PATH, payload.model_dump())
    b64 = base64.b64encode(docx_bytes).decode("utf-8")
    return {"filename": "FORMULARIO_CM_PRENCHIDO.docx", "filedata": b64}

@app.post("/fill_url")
def fill_url(request: Request, payload: CMPayload, x_api_key: Optional[str] = Header(default=None)):
    """
    Salva o DOCX no servidor e retorna a URL pública para download.
    """
    _check_api_key(x_api_key)
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template não encontrado: {TEMPLATE_PATH}")

    # gera o arquivo
    docx_bytes = preencher_docx_from_payload(TEMPLATE_PATH, payload.model_dump())

    # nome único
    stamp = datetime.utcnow().strftime("%Y%m%d-%H%M%S")
    unique = uuid.uuid4().hex[:8]
    filename = f"FORMULARIO_CM_PRENCHIDO_{stamp}_{unique}.docx"
    filepath = os.path.join(DOWNLOAD_DIR, filename)

    # grava em disco
    with open(filepath, "wb") as f:
        f.write(docx_bytes)

    # URL pública
    base_url = str(request.base_url).rstrip("/")
    file_url = f"{base_url}/downloads/{filename}"

    return {"filename": filename, "fileUrl": file_url}
