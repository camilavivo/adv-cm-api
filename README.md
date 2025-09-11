# ADVFARMA CM Filler API

API em FastAPI que preenche o formulário oficial de Controle de Mudanças (CM) no modelo Word (.docx).

## Uso local
```bash
pip install -r requirements.txt
uvicorn app:app --reload
```

## Endpoints
- `GET /health` → status OK
- `POST /fill` → recebe JSON e devolve DOCX preenchido

Headers: `x-api-key: SUA_CHAVE`
