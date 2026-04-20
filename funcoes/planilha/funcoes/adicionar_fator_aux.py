from openpyxl.utils import column_index_from_string
from funcoes.common.copiar_coluna import copiar_coluna_com_numeros
from funcoes.common.valor_bdi_final import valor_bdi_final
from funcoes.get.get_linhas_json import *


def copiar_colunas(sheet, dados):
    copiar_coluna_com_numeros(
        sheet, get_coeficiente_aux(dados), get_copiar_coeficiente_aux(dados)
    )
    copiar_coluna_com_numeros(
        sheet, get_preco_unitario_aux(dados), get_copiar_preco_unitario_aux(dados)
    )


def adicionar_formula_preco_unitario_menos_preco_antigo(sheet, dados):
    origem = get_copiar_preco_unitario_aux(dados)
    destino_idx = column_index_from_string(origem) + 1
    preco_unit = get_preco_unitario_aux(dados)

    for i, cell in enumerate(sheet[origem], start=1):
        if cell.value is not None:
            sheet.cell(row=i, column=destino_idx).value = (
                f"=({origem}{i}-{preco_unit}{i})"
            )


def adicionar_fator_aux(workbook, dados):
    sheet_aux = workbook[get_planilha_aux(dados)]
    copiar_colunas(sheet_aux, dados)
    adicionar_formula_preco_unitario_menos_preco_antigo(sheet_aux, dados)
    valor_bdi_final(
        sheet_aux, dados, get_coluna_totais_aux(dados), get_valor_totais_aux(dados)
    )
