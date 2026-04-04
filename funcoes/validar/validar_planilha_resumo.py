"""
Componente para validar a planilha RESUMO (FATOR) com correção interativa.
"""

from funcoes.get.get_linhas_json import get_valor_total_resumo_string, get_coluna_total_resumo
from funcoes.validar.validar_celula_bdi import validar_celula_bdi
from funcoes.validar.validar_coluna import validar_coluna_existe
from funcoes.validar.validar_nome_planilha import validar_nome_planilha
from funcoes.validar.validar_valor_existe_na_coluna import validar_valor_existe_na_coluna


def validar_planilha_resumo(workbook, dados, erros, indice_config=0):
    """
    Valida a planilha RESUMO (FATOR) com correção interativa.

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros
        indice_config: Índice da configuração no JSON

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None)
    """
    nome_planilha = dados[indice_config].get("planilhaFator", "RESUMO")

    valido, sheet, _ = validar_nome_planilha(
        workbook,
        nome_planilha,
        "Resumo",
        erros,
        dados,
        indice_config,
        workbook.sheetnames if hasattr(workbook, "sheetnames") else [],
    )
    if not valido:
        return False, None

    if sheet and sheet.max_row < 2:
        erros.append(
            f"ERRO: A aba '{nome_planilha}' está vazia ou não tem dados suficientes."
        )
        return False, sheet

    coluna_fator = dados[indice_config].get("colunaFator", "G")
    linha_fator = dados[indice_config].get("linhaFator", "4")

    if sheet:
        if not validar_coluna_existe(
            sheet,
            coluna_fator,
            "Fator",
            erros,
            dados,
            indice_config,
            nome_planilha,
            "colunaFator",
        ):
            return False, sheet

        bdi_valido, _, _ = validar_celula_bdi(
            sheet, coluna_fator, linha_fator, erros, dados, indice_config, nome_planilha
        )
        if not bdi_valido:
            return False, sheet

        valor_total_resumo = dados[indice_config].get(
            "valorTotalResumo", "VALOR TOTAL RESUMO:"
        )
        coluna_total_resumo = dados[indice_config].get("colunaTotalResumo", "C")

        if valor_total_resumo:
            if not validar_valor_existe_na_coluna(
                sheet,
                coluna_total_resumo,
                valor_total_resumo,
                "Valor Total do Resumo",
                nome_planilha,
                erros,
                dados,
                indice_config,
                "valorTotalResumo",
            ):
                return False, sheet

    return True, sheet
