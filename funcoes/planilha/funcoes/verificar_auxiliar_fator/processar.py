from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell

from .constantes import TEXTOS_SKIP, TEXTOS_VALOR_SKIP, VALOR_LABEL
from .hyperlink import _add_hyperlink
from .limpar import _codigo_valido, _limpar_codigo


def processar_planilha(
    sheet,
    col_desc,
    col_coef,
    col_preco,
    col_coef_antigo,
    col_preco_antigo,
    mapa_nome_inicia,
    mapa_config,
    planilha_destino,
    mapa_titulos_aux,
    is_auxiliar=False,
):
    """FOR único percorrendo linhas, identificando seções e aplicando regras.

    Identifica:
    - INÍCIO: quando encontra nome da seção (ex: "SERVIÇOS")
    - FIM: quando encontra total da seção (ex: "TOTAL SERVIÇOS:")

    Aplica:
    - adicionarFator: "Sim" → adiciona fórmula com fator
    - buscarAuxiliar: "Sim" → cria hyperlink e fórmula referencing planilha auxiliar
    """
    resultado = {
        "formulas_fator": 0,
        "formulas_auxiliar": 0,
        "hyperlinks": 0,
    }

    max_row = min(sheet.max_row, 20000)

    # Estado da seção atual
    secao_atual = None
    linha_inicio_secao = None
    linha_fim_secao = None
    config_secao = None

    # FOR único percorrendo todas as linhas
    for linha in range(1, max_row + 1):
        cell_desc = sheet.cell(row=linha, column=col_desc)

        if isinstance(cell_desc, MergedCell):
            continue

        valor = cell_desc.value
        if not valor:
            continue

        valor_str = str(valor).replace("\u200b", "").replace("\ufeff", "").strip()
        valor_upper = valor_str.upper()

        # ==========================================
        # IDENTIFICAR INÍCIO DE SEÇÃO (pelo nome)
        # ==========================================
        if secao_atual is None:
            for config in mapa_config:
                if valor_upper == config["nome"].upper():
                    # Encontrou início de nova seção
                    secao_atual = config["nome"]
                    linha_inicio_secao = linha
                    config_secao = config

                    # Se for auxiliar, construir mapa de códigos
                    if is_auxiliar:
                        codigo = _limpar_codigo(valor_str)
                        if codigo and len(codigo) >= 5:
                            # Buscar linha com VALOR: abaixo
                            linha_valor = buscar_linha_valor(sheet, linha, max_row)
                            if linha_valor:
                                mapa_titulos_aux[codigo.upper()] = linha_valor
                    break

        # ==========================================
        # IDENTIFICAR FIM DE SEÇÃO (pelo total)
        # ==========================================
        if secao_atual and config_secao:
            total_upper = config_secao.get("total", "").upper()
            if valor_upper == total_upper:
                linha_fim_secao = linha
                # Reset para próxima seção
                secao_atual = None
                linha_inicio_secao = None
                config_secao = None
                continue

        # ==========================================
        # PROCESSAR LINHAS DENTRO DA SEÇÃO
        # ==========================================
        if secao_atual and config_secao and linha > linha_inicio_secao:
            # Verificar se não mudou de seção
            if linha_fim_secao and linha >= linha_fim_secao:
                secao_atual = None
                linha_inicio_secao = None
                config_secao = None
                continue

            # Pular linhas com TEXTOS_SKIP
            if any(x in valor_upper for x in TEXTOS_SKIP):
                continue
            if any(x in valor_upper for x in TEXTOS_VALOR_SKIP):
                continue

            # Pular linhas de cabeçalho
            if "COEFICIENTE" in valor_upper or "PREÇO UNITÁRIO" in valor_upper:
                continue

            # Verificar fonte/unidade
            cell_fonte = sheet.cell(row=linha, column=col_desc + 2).value
            cell_unid = sheet.cell(row=linha, column=col_desc + 3).value
            if cell_fonte and "FONTE" in str(cell_fonte).upper():
                continue
            if cell_unid and "UNID" in str(cell_unid).upper():
                continue

            # Verificar se é código válido
            if not _codigo_valido(sheet, linha, col_desc):
                continue

            # Verificar filtros (iniciaPor, naoIniciaPor)
            iniciaPor = ""
            naoIniciaPor = ""
            for item in mapa_nome_inicia:
                if item["nome"].upper() == secao_atual.upper():
                    iniciaPor = item.get("iniciaPor", "")
                    naoIniciaPor = item.get("naoIniciaPor", "")
                    break

            codigo_upper = valor_str.upper()
            if iniciaPor and not codigo_upper.startswith(iniciaPor.upper()):
                continue
            if naoIniciaPor and codigo_upper.startswith(naoIniciaPor.upper()):
                continue

            # ==========================================
            # ADICIONAR FATOR
            # ==========================================
            resultado["formulas_fator"] += adicionar_fator(
                sheet,
                linha,
                col_coef,
                col_preco,
                col_coef_antigo,
                col_preco_antigo,
                config_secao,
                secao_atual,
            )

            # ==========================================
            # BUSCAR AUXILIAR
            # ==========================================
            if config_secao.get("buscarAuxiliar") == "Sim":
                resultado_busca = buscar_auxiliar(
                    sheet,
                    linha,
                    col_desc,
                    col_preco,
                    valor_str,
                    planilha_destino,
                    mapa_titulos_aux,
                    is_auxiliar,
                )
                resultado["hyperlinks"] += resultado_busca["hyperlinks"]
                resultado["formulas_auxiliar"] += resultado_busca["formulas_auxiliar"]

    return resultado


def adicionar_fator(
    sheet,
    linha,
    col_coef,
    col_preco,
    col_coef_antigo,
    col_preco_antigo,
    config_secao,
    secao_atual,
):
    """Adiciona fórmula de fator na linha."""
    count = 0
    if config_secao.get("adicionarFator") == "Sim":
        val_coef = sheet.cell(row=linha, column=col_coef).value
        val_preco = sheet.cell(row=linha, column=col_preco).value

        # Pular se já tem *FATOR
        if (
            val_coef and isinstance(val_coef, str) and "*FATOR" in val_coef.upper()
        ) or (
            val_preco and isinstance(val_preco, str) and "*FATOR" in val_preco.upper()
        ):
            return 0

        if config_secao.get("fatorCoeficiente"):
            cell = sheet.cell(row=linha, column=col_coef)
            if not isinstance(cell, MergedCell):
                cell.value = f"={get_column_letter(col_coef_antigo)}{linha}*FATOR"
                count = 1
        else:
            cell = sheet.cell(row=linha, column=col_preco)
            if not isinstance(cell, MergedCell):
                cell.value = (
                    f"=ROUND({get_column_letter(col_preco_antigo)}{linha}*FATOR, 2)"
                )
                count = 1

    return count


def buscar_auxiliar(
    sheet,
    linha,
    col_desc,
    col_preco,
    valor_str,
    planilha_destino,
    mapa_titulos_aux,
    is_auxiliar,
):
    """Busca código no mapa e cria hyperlink/fórmula."""
    resultado = {"hyperlinks": 0, "formulas_auxiliar": 0}

    cell_desc = sheet.cell(row=linha, column=col_desc)
    if cell_desc.hyperlink:
        return resultado

    codigo_limpo = _limpar_codigo(valor_str)
    if codigo_limpo and len(codigo_limpo) >= 5:
        codigo_upper = codigo_limpo.upper()

        if codigo_upper in mapa_titulos_aux:
            linha_valor = mapa_titulos_aux[codigo_upper]

            if linha_valor:
                # Criar hyperlink
                _add_hyperlink(sheet, linha, col_desc, planilha_destino, linha_valor)
                resultado["hyperlinks"] = 1

                # Adicionar fórmula no preço unitário
                cell_preco = sheet.cell(row=linha, column=col_preco)
                if not isinstance(cell_preco, MergedCell):
                    if is_auxiliar:
                        if linha_valor > linha:
                            cell_preco.value = f"=G{linha_valor}"
                            resultado["formulas_auxiliar"] = 1
                    else:
                        cell_preco.value = f"='{planilha_destino}'!G{linha_valor}"
                        resultado["formulas_auxiliar"] = 1

    return resultado


def buscar_linha_valor(sheet, linha_inicio, max_row):
    """Busca linha com 'VALOR:' a partir da linha_inicio."""
    for linha in range(linha_inicio + 1, max_row + 1):
        val_e = sheet.cell(row=linha, column=5).value
        if val_e and VALOR_LABEL.upper() in str(val_e).upper():
            return linha
    return None
