# -*- coding: utf-8 -*-
from docx import Document

LABELS = {
    "DATA": "DATA",
    "SOLICITANTE": "SOLICITANTE",
    "DEPARTAMENTO": "DEPARTAMENTO",
    "TITULO": "TÍTULO MUDANÇA",
    "CARATER": "CARÁTER MUDANÇA",
    "RETORNO": "RETORNO MUDANÇA TEMPORÁRIA",
    "SITUACAO": "SITUAÇÃO ATUAL",
    "ALTERACAO": "ALTERAÇÃO PROPOSTA",
    "JUST_MUD": "JUSTIFICATIVA MUDANÇA",
    "DESC_ITEM": "DESCRIÇÃO ITEM",
    "NUM_CORR": "NÚMERO CORRESPONDENTE",
    "ABRANGENCIA": "ABRANGÊNCIA DA MUDANÇA",
    "REFERE": "MUDANÇA REFERE-SE Á",
    "IMPACTO": "POTENCIAL IMPACTO AVALIADO",
    "CLASSIF": "CLASSIFICAÇÃO DA CRITICIDADE",
    "JUST_CLASSIF": "JUSTIFICATIVA DA CLASSIFICAÇÃO",
    "ANEXOS": "ANEXOS:",
    "PLANO": "PLANO DE IMPLEMENTAÇÃO",
    "TREIN": "EXECUÇÃO DO TREINAMENTO",
    "VOE": "VERIFICAÇÃO DE EFICÁCIA (VoE) PÓS IMPLEMENTAÇÃO",
    "RES_VOE": "RESULTADOS DA VoE",
    "OBS": "Observações finais",
}

TODOS_SETORES_MODELO = [
    "Farmacêutico Responsável","Garantia da Qualidade","Operações","Diretoria ou Conselho",
    "Almoxarifado","Comercial","Compras","Controle da Qualidade","Expedição","Faturamento",
    "Financeiro / Custos","Fiscal/Contábil","Informática (TI)","Manutenção","Planejamento (PCP)",
    "Produção","Regulatório","Terceiros","Validação",
]
OBRIGATORIOS_SECAO5 = {"Farmacêutico Responsável","Diretoria ou Conselho","Regulatório"}

def _set_cell_right_of_label(doc: Document, label: str, value: str) -> bool:
    lab_upper = label.strip().upper()
    for table in doc.tables:
        for row in table.rows:
            left = row.cells[0].text.strip().upper()
            if lab_upper == left or lab_upper in left:
                row.cells[1].text = value
                return True
    return False

def _preencher_secao5(doc: Document, departamentos_pertinentes):
    pertinentes = set(OBRIGATORIOS_SECAO5 | set(departamentos_pertinentes or []))
    for table in doc.tables:
        for row in table.rows:
            dept = row.cells[0].text.strip()
            if dept in TODOS_SETORES_MODELO and dept not in pertinentes:
                if len(row.cells) >= 3:
                    row.cells[1].text = "Não aplicável"
                    row.cells[2].text = "Não aplicável"

def preencher_docx_from_payload(template_path: str, payload: dict) -> bytes:
    doc = Document(template_path)
    join = lambda xs: "\n".join(xs) if xs else "—"

    _set_cell_right_of_label(doc, LABELS["DATA"], payload["data"])
    _set_cell_right_of_label(doc, LABELS["SOLICITANTE"], payload["solicitante"])
    _set_cell_right_of_label(doc, LABELS["DEPARTAMENTO"], payload["departamento"])
    _set_cell_right_of_label(doc, LABELS["TITULO"], payload.get("titulo_mudanca") or "—")
    _set_cell_right_of_label(doc, LABELS["CARATER"], payload.get("caracter_mudanca","Temporária"))
    _set_cell_right_of_label(doc, LABELS["RETORNO"], payload.get("retorno_mudanca_temp") or "—")
    _set_cell_right_of_label(doc, LABELS["SITUACAO"], payload["situacao_atual"])
    _set_cell_right_of_label(doc, LABELS["ALTERACAO"], payload["alteracao_proposta"])
    _set_cell_right_of_label(doc, LABELS["JUST_MUD"], "Problema → impacto → mitigação descritos na prévia.")
    _set_cell_right_of_label(doc, LABELS["DESC_ITEM"], join(payload.get("descricoes_itens")))
    _set_cell_right_of_label(doc, LABELS["NUM_CORR"], join(payload.get("numeros_correspondentes")))
    _set_cell_right_of_label(doc, LABELS["ABRANGENCIA"], payload["abrangencia"])
    _set_cell_right_of_label(doc, LABELS["REFERE"], join(payload.get("mudanca_refere_se")))
    _set_cell_right_of_label(doc, LABELS["IMPACTO"], join(payload.get("impactos")))
    _set_cell_right_of_label(doc, LABELS["CLASSIF"], payload["classificacao"])
    _set_cell_right_of_label(doc, LABELS["JUST_CLASSIF"], payload["justificativa_classificacao"])
    _set_cell_right_of_label(doc, LABELS["ANEXOS"], join(payload.get("anexos_aplicaveis")))
    _preencher_secao5(doc, payload.get("departamentos_pertinentes"))
    _set_cell_right_of_label(doc, LABELS["PLANO"], join(payload.get("plano_implementacao")))
    if "treinamento_executado" in payload and payload["treinamento_executado"] is not None:
        _set_cell_right_of_label(doc, LABELS["TREIN"], "Sim" if payload["treinamento_executado"] else "Não")
    _set_cell_right_of_label(doc, LABELS["VOE"],
        f"Critérios: {payload.get('voe_criterios','—')}\nPeríodo: {payload.get('voe_periodo','—')}")
    _set_cell_right_of_label(doc, LABELS["RES_VOE"], payload.get("voe_resultados_esperados","—"))
    _set_cell_right_of_label(doc, LABELS["OBS"], "Encerrar após VoE e retorno/restabelecimento da condição original.")

    from io import BytesIO
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
