from openpyxl.utils import column_index_from_string
from funcoes.common.copiar_coluna import copiar_coluna_com_numeros
from funcoes.common.valor_bdi_final import valor_bdi_final
from funcoes.get.get_linhas_json import *


def criar_link_composicao_orcamentaria(
    sheet_origem,
    col_desc,
    linha_origem,
    nome_planilha_destino,
    col_valor,
    linha_destino,
):
    celula = sheet_origem[f"{col_desc}{linha_origem}"]
    celula.hyperlink = f"#{nome_planilha_destino}!{col_valor}{linha_destino}"


def copiar_colunas(sheet, dados):
    copiar_coluna_com_numeros(
        sheet, get_coeficiente_comp(dados), get_copiar_coeficiente_comp(dados)
    )
    copiar_coluna_com_numeros(
        sheet, get_preco_unitario_comp(dados), get_copiar_preco_unitario_comp(dados)
    )


def adicionar_formula_preco_unitario_menos_preco_antigo(sheet, dados):
    origem = get_copiar_preco_unitario_comp(dados)
    destino = column_index_from_string(origem) + 1
    preco_unit = get_preco_unitario_comp(dados)

    for i, cell in enumerate(sheet[origem], start=1):
        if cell.value is not None:
            sheet.cell(row=i, column=destino).value = f"=({origem}{i}-{preco_unit}{i})"


def adicionar_fator_comp(workbook, dados):
    sheet_comp = workbook[get_planilha_comp(dados)]
    copiar_colunas(sheet_comp, dados)
    adicionar_formula_preco_unitario_menos_preco_antigo(sheet_comp, dados)
    valor_bdi_final(
        sheet_comp, dados, get_coluna_totais_comp(dados), get_valor_totais_comp(dados)
    )
    return True, None
