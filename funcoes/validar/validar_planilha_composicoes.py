"""
Componente para validar a planilha COMPOSIÇÕES com correção interativa.
"""

from funcoes.get.get_linhas_json import (
    get_coluna_totais_comp,
    get_valor_totais_comp,
    get_valor_com_bdi_string,
    get_valor_bdi_comp,
    get_valor_string,
)
from funcoes.validar.validar_coluna import validar_coluna_existe
from funcoes.validar.validar_nome_planilha import validar_nome_planilha
from funcoes.validar.validar_valor_existe_na_coluna import validar_valor_existe_na_coluna


def validar_planilha_composicoes(workbook, dados, erros, indice_config=0):
    """
    Valida a planilha COMPOSIÇÕES com correção interativa.

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros
        indice_config: Índice da configuração no JSON

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None)
    """
    nome_planilha = dados[indice_config].get("planilhaComposicao", "COMPOSICOES")

    valido, sheet, _ = validar_nome_planilha(
        workbook,
        nome_planilha,
        "Composições",
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

    if sheet:
        colunas_comp = [
            ("composicaoDescricao", "Descrição", "A"),
            ("colunaItemDescricaoComposicao", "Código do Item", "B"),
            ("composicaoCoeficiente", "Coeficiente", "E"),
            ("composicaoPrecoUnitario", "Preço Unitário", "F"),
            ("composicaoCoeficienteCopiar", "Coeficiente (copiar)", "L"),
            ("composicaoPrecoUnitarioCopiar", "Preço Unitário (copiar)", "M"),
            ("colunaTotaisComposicao", "Coluna de Totais", "E"),
        ]

        for campo_json, nome_col, default_col in colunas_comp:
            col = dados[indice_config].get(campo_json, default_col)
            if not validar_coluna_existe(
                sheet,
                col,
                nome_col,
                erros,
                dados,
                indice_config,
                nome_planilha,
                campo_json,
            ):
                return False, sheet

        col_totais = get_coluna_totais_comp(dados[indice_config])
        col_valor_totais = get_valor_totais_comp(dados[indice_config])

        if not validar_coluna_existe(
            sheet,
            col_totais,
            "Coluna de Totais",
            erros,
            dados,
            indice_config,
            nome_planilha,
            "colunaTotaisComposicao",
        ):
            return False, sheet

        if not validar_coluna_existe(
            sheet,
            col_valor_totais,
            "Coluna de Valores Totais",
            erros,
            dados,
            indice_config,
            nome_planilha,
            "valorTotaisComposicao",
        ):
            return False, sheet

        valores_a_verificar = {
            "total_com_bdi": (
                get_valor_com_bdi_string(dados[indice_config]),
                "valorComBdi",
            ),
            "valor_bdi": (get_valor_bdi_comp(dados[indice_config]), "valorBdi"),
            "valor_string": (get_valor_string(dados[indice_config]), "valor"),
        }

        for nome_valor, (valor_buscado, campo_json) in valores_a_verificar.items():
            if not validar_valor_existe_na_coluna(
                sheet,
                col_totais,
                valor_buscado,
                nome_valor,
                nome_planilha,
                erros,
                dados,
                indice_config,
                campo_json,
            ):
                return False, sheet

    return True, sheet
