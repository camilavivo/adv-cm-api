# -*- coding: utf-8 -*-
import os
from fastapi import FastAPI, HTTPException, Header, Response
from pydantic import BaseModel, Field
from typing import List, Optional
from cm_filler import preencher_docx_from_payload

API_KEY = os.getenv("CM_API_KEY")
TEMPLATE_PATH = os.getenv("CM_TEMPLATE_PATH", "TESTE FORMULÁRIO CM.docx")

app = FastAPI(title="ADV CM Filler API",
              version="1.0.0",
              description="API que preenche o formulário CM e retorna o DOCX.")

class CMPayload(BaseModel):
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
    anexos_aplicaveis: List[str] = Field(default_factory=list)
    departamentos_pertinentes: List[str] = Field(default_factory=list)
    treinamento_executado: Optional[bool] = None
    plano_implementacao: List[str] = Field(default_factory=list)
    voe_criterios: Optional[str] = ""
    voe_periodo: Optional[str] = ""
    voe_resultados_esperados: Optional[str] = ""

def _check_api_key(x_api_key: str | None):
    if not API_KEY:
        return
    if not x_api_key or x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key inválida")

@app.get("/health")
def health():
    return {"status": "ok"}

@app.post("/fill", response_class=Response,
          responses={200: {"content": {"application/vnd.openxmlformats-officedocument.wordprocessingml.document": {}}}})
def fill(payload: CMPayload, x_api_key: Optional[str] = Header(default=None)):
    _check_api_key(x_api_key)
    import os
    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(status_code=500, detail=f"Template não encontrado: {TEMPLATE_PATH}")
    docx_bytes = preencher_docx_from_payload(TEMPLATE_PATH, payload.model_dump())
    headers = {"Content-Disposition": 'attachment; filename="FORMULARIO_CM_PRENCHIDO.docx"'}
    return Response(content=docx_bytes,
                    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    headers=headers)
