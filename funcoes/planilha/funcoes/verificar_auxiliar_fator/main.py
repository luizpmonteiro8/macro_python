from openpyxl.utils import column_index_from_string

from .config import extrair_configuracoes
from .mapa_mescladas import construir_mapa_mescladas
from .processar import processar_planilha


def verificar_auxiliar_fator(workbook, dados, todos_item):
    """Processa composições e auxiliares.

    Fluxo:
    1. Extrai config do valores_item.json
    2. Processa COMPOSICOES AUXILIARES primeiro (constrói mapa de códigos)
    3. Processa COMPOSICOES depois (usa mapa para hyperlinks)
    """
    dados_itens = dados[0] if isinstance(dados, list) else dados

    # Extrair config do JSON
    mapa_nome_inicia, mapa_config = extrair_configuracoes(todos_item)

    # Configurações das planilhas
    planilha_comp = dados_itens.get("planilhaComposicao", "COMPOSICOES")
    planilha_aux = dados_itens.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")

    col_desc = column_index_from_string(dados_itens.get("composicaoDescricao", "A"))
    col_coef = column_index_from_string(dados_itens.get("composicaoCoeficiente", "E"))
    col_preco = column_index_from_string(
        dados_itens.get("composicaoPrecoUnitario", "F")
    )
    col_coef_antigo = column_index_from_string(
        dados_itens.get("composicaoCoeficienteCopiar", "L")
    )
    col_preco_antigo = column_index_from_string(
        dados_itens.get("composicaoPrecoUnitarioCopiar", "M")
    )

    # Obter worksheets
    sheet_comp = workbook[planilha_comp]
    sheet_aux = workbook[planilha_aux]

    # Mapa de códigos - construído ANTES de processar_planilha usando células mescladas
    mapa_titulos_aux = construir_mapa_mescladas(sheet_aux, col_desc)

    # Processar COMPOSICOES AUXILIARES primeiro (já tem mapa pré-construído)
    resultado_aux = processar_planilha(
        sheet_aux,
        col_desc,
        col_coef,
        col_preco,
        col_coef_antigo,
        col_preco_antigo,
        mapa_nome_inicia,
        mapa_config,
        planilha_aux,
        mapa_titulos_aux,
        is_auxiliar=True,
    )

    # Processar COMPOSICOES depois
    resultado_comp = processar_planilha(
        sheet_comp,
        col_desc,
        col_coef,
        col_preco,
        col_coef_antigo,
        col_preco_antigo,
        mapa_nome_inicia,
        mapa_config,
        planilha_aux,
        mapa_titulos_aux,
        is_auxiliar=False,
    )

    return {
        "formulas_fator_comp": resultado_comp["formulas_fator"],
        "formulas_fator_aux": resultado_aux["formulas_fator"],
        "formulas_auxiliares_comp": resultado_comp["formulas_auxiliar"],
        "formulas_auxiliares_aux": resultado_aux["formulas_auxiliar"],
        "hyperlinks_criados": resultado_comp["hyperlinks"]
        + resultado_aux["hyperlinks"],
    }
