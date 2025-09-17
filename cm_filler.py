# -*- coding: utf-8 -*-
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ====== CONFIG PADRÃO DE FORMATAÇÃO ======
FONT_NAME = "Arial"
FONT_SIZE_PT = 9
FONT_COLOR = RGBColor(89, 89, 89)  # Branco (Plano de Fundo 1), mais escuro 35%

def _format_paragraph(p):
    """Aplica alinhamento e fonte padrão ao parágrafo e seus runs."""
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    # se não houver runs, cria um vazio para aplicar formato
    if not p.runs:
        r = p.add_run("")
        r.font.name = FONT_NAME
        r.font.size = Pt(FONT_SIZE_PT)
        r.font.color.rgb = FONT_COLOR
        return
    for r in p.runs:
        r.font.name = FONT_NAME
        r.font.size = Pt(FONT_SIZE_PT)
        r.font.color.rgb = FONT_COLOR

def _write_cell_value(cell, value: str):
    """
    Escreve na célula preservando nosso padrão:
    - quebra value por '\n' e cria um parágrafo para cada linha
    - aplica Arial 9, justificado, cor RGB(89,89,89)
    """
    if value is None:
        value = "—"
    # limpa conteúdo mantendo a própria célula
    # (cell.text = "" recria um parágrafo padrão; vamos formatá-lo)
    cell.text = ""
    lines = str(value).split("\n")
    # o primeiro parágrafo já existe
    if lines:
        cell.paragraphs[0].runs[0].text = lines[0] if cell.paragraphs[0].runs else lines[0]
        _format_paragraph(cell.paragraphs[0])
        # linhas adicionais viram novos parágrafos
        for ln in lines[1:]:
            p = cell.add_paragraph(ln)
            _format_paragraph(p)
    else:
        # valor vazio — ainda assim formatar o parágrafo
        _format_paragraph(cell.paragraphs[0])

LABELS = {
    "DATA": "DATA",
    "SOLICITANTE": "SOLICITANTE",
    "DEPARTAMENTO": "DEPARTAMENTO",
    "TITULO": "TÍTULO MUDANÇA",
    "CARATER": "CARÁTER MUDANÇA",
    "RETORNO": "RETORNO MUDANÇA",             # Rev.13
    "SITUACAO": "SITUAÇÃO ATUAL",
    "ALTERACAO": "ALTERAÇÃO PROPOSTA",
    "JUST_MUD": "JUSTIFICATIVA MUDANÇA",      # Rev.13
    "DESC_ITEM": "DESCRIÇÃO ITEM",
    "NUM_CORR": "NÚMERO CORRESPONDENTE",
    "ABRANGENCIA": "ABRANGÊNCIA DA MUDANÇA",
    "REFERE": "MUDANÇA REFERE-SE Á",
    "IMPACTO": "POTENCIAL IMPACTO AVALIADO",
    "CLASSIF": "CLASSIFICAÇÃO DA CRITICIDADE",
    "JUST_CLASSIF": "JUSTIFICATIVA DA CLASSIFICAÇÃO",
    "ANEXOS": "ANEXOS",                        # Rev.13
    "PLANO": "PLANO DE IMPLEMENTAÇÃO",
    "TREIN": "EXECUÇÃO DO TREINAMENTO",
    "VOE": "VERIFICAÇÃO DE EFICÁCIA (VoE) PÓS IMPLEMENTAÇÃO",
    "RES_VOE": "RESULTADOS DA VoE",
    "OBS": "OBSERVAÇÕES FINAIS",              # Rev.13
}

TODOS_SETORES_MODELO = [
    "Farmacêutico Responsável","Garantia da Qualidade","Operações","Diretoria ou Conselho",
    "Almoxarifado","Comercial","Compras","Controle da Qualidade","Expedição","Faturamento",
    "Financeiro / Custos","Fiscal/Contábil","Informática (TI)","Manutenção","Planejamento (PCP)",
    "Produção","Regulatório","Terceiros","Validação",
]
OBRIGATORIOS_SECAO5 = {"Farmacêutico Responsável","Diretoria ou Conselho","Regulatório"}

def _set_cell_right_of_label(doc: Document, label: str, value: str) -> bool:
    """Localiza a linha cujo 1º campo corresponde ao label e escreve no 2º campo com formatação padrão."""
    lab_upper = label.strip().upper()
    for table in doc.tables:
        for row in table.rows:
            left = row.cells[0].text.strip().upper()
            if lab_upper == left or lab_upper in left:
                _write_cell_value(row.cells[1], value)
                return True
    return False

def _preencher_secao5(doc: Document, departamentos_pertinentes):
    """Marca 'Não aplicável' para departamentos não pertinentes na Seção 5, aplicando formatação padrão."""
    pertinentes = set(OBRIGATORIOS_SECAO5 | set(departamentos_pertinentes or []))
    for table in doc.tables:
        for row in table.rows:
            dept = row.cells[0].text.strip()
            if dept in TODOS_SETORES_MODELO and dept not in pertinentes:
                if len(row.cells) >= 3:
                    _write_cell_value(row.cells[1], "Não aplicável")
                    _write_cell_value(row.cells[2], "Não aplicável")

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
    # Justificativa da mudança — agora vem do payload
    _set_cell_right_of_label(doc, LABELS["JUST_MUD"], payload.get("justificativa_mudanca", "—"))
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
    _set_cell_right_of_label(doc, LABELS["OBS"], payload.get("observacoes_finais","Encerrar após VoE e retorno/restabelecimento da condição original."))

    from io import BytesIO
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()
