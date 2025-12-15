import tkinter as tk
from openpyxl.utils import column_index_from_string

from funcoes.common.buscar_palavras import (
    buscar_palavra_com_linha,
    buscar_palavra_com_linha_exato,
    buscar_palavra_contem,
)
from funcoes.common.copiar_coluna import copiar_coluna_com_numeros
from funcoes.common.valor_bdi_final import valor_bdi_final
from funcoes.get.get_linhas_json import *
from funcoes.planilha.funcoes.adicionar_fator_aux import adicionar_fator_totais_aux


def criar_link_composicao(
    sheet_origem,
    col_desc,
    linha_origem,
    nome_planilha_destino,
    col_valor,
    linha_destino,
):
    """
    Cria um hyperlink na célula de descrição apontando para o total da composição
    """
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


def fator_nos_item_totais(
    sheet,
    dados,
    lin_ini,
    lin_fim,
    nome,
    totalNome,
    coeficiente,
    adicionar_fator,
    inicia_por=None,
    nao_inicia_por=None,
):
    col_desc = get_descricao_comp(dados)
    col_totais = get_coluna_totais_comp(dados)
    col_valor = get_valor_totais_comp(dados)
    col_preco = get_preco_unitario_comp(dados)
    col_coef = get_coeficiente_comp(dados)
    col_preco_antigo = get_copiar_preco_unitario_comp(dados)
    col_coef_antigo = get_copiar_coeficiente_comp(dados)

    inicial = buscar_palavra_com_linha_exato(sheet, col_desc, nome, lin_ini, lin_fim)
    final = buscar_palavra_com_linha_exato(
        sheet, col_totais, totalNome, lin_ini, lin_fim
    )

    if -1 < inicial < final:
        # soma final
        sheet[f"{col_valor}{final}"].value = (
            f"=SUM({col_valor}{inicial+1}:{col_valor}{final-1})"
        )

        for y in range(inicial + 1, final):
            desc = sheet[f"{col_desc}{y}"].value
            if inicia_por and (desc is None or not desc.startswith(inicia_por)):
                continue
            if nao_inicia_por and (desc is None or desc.startswith(nao_inicia_por)):
                continue

            if coeficiente and adicionar_fator:
                sheet[f"{col_coef}{y}"].value = f"={col_coef_antigo}{y}*FATOR"
            elif adicionar_fator:
                sheet[f"{col_preco}{y}"].value = (
                    f"=ROUND({col_preco_antigo}{y}*FATOR, 2)"
                )

        return inicial, final


def buscar_comp_auxiliar(workbook, dados, itemChave, lin, lin_total):
    sheet_comp = workbook[get_planilha_comp(dados)]
    sheet_aux = workbook[get_planilha_aux(dados)]

    col_item = get_item_descricao_comp_aux(dados)
    col_desc_aux = get_descricao_aux(dados)
    col_totais_aux = get_coluna_totais_aux(dados)
    col_valor_aux = get_valor_totais_aux(dados)
    col_preco_comp = get_preco_unitario_comp(dados)
    valor_string = get_valor_string(dados)

    ultima_busca = 1
    ultima_linha_aux = sheet_aux.max_row

    for x in range(lin, lin_total):
        cod = sheet_comp[f"{col_desc_aux}{x}"].value
        nome = sheet_comp[f"{col_item}{x}"].value
        if nome is None:
            continue

        # primeira tentativa busca completa
        linha_ini = buscar_palavra_com_linha(
            sheet_aux, col_desc_aux, f"{cod} {nome}", ultima_busca, ultima_linha_aux
        )
        if linha_ini == -1:
            linha_ini = buscar_palavra_com_linha(
                sheet_aux, col_desc_aux, cod, ultima_busca, ultima_linha_aux
            )

        if cod in ("I0690", "I0769"):
            print(
                f"{cod} {nome} -> linha_ini: {linha_ini}, ultima_linha_aux: {ultima_linha_aux}"
            )

        if linha_ini > -1:
            linha_fim = buscar_palavra_com_linha(
                sheet_aux, col_totais_aux, valor_string, linha_ini, ultima_linha_aux
            )
            if cod in ("I0690", "I0769"):
                print(
                    f"final: {linha_fim}, valor_string: {valor_string}, linha_inicial: {linha_ini}"
                )

            adicionar_fator_totais_aux(workbook, dados, itemChave, linha_ini, linha_fim)
            sheet_comp[f"{col_preco_comp}{x}"].value = (
                f"='{get_planilha_aux(dados)}'!{col_valor_aux}{linha_fim}"
            )


def adicionar_fator_totais(workbook, dados, itemChave, lin_ini, lin_fim):
    sheet = workbook[get_planilha_orcamentaria(dados)]
    sheet_comp = workbook[get_planilha_comp(dados)]
    sheet_comp_max = sheet_comp.max_row + 1

    col_preco_planilha = get_planilha_preco_unitario(dados)
    col_desc_comp = get_descricao_comp(dados)
    col_totais_comp = get_coluna_totais_comp(dados)
    col_valor_comp = get_valor_totais_comp(dados)
    col_cod = get_planilha_codigo(dados)
    col_desc = get_planilha_descricao(dados)
    valor_com_bdi = get_valor_com_bdi_string(dados)
    valor_string = get_valor_string(dados)

    itens_array = [v for k, v in itemChave.items() if k.startswith("item")]
    linha_busca_ini = 1

    for x in range(lin_ini, lin_fim):
        cod = sheet[f"{col_cod}{x}"].value
        descricao = sheet[f"{col_desc}{x}"].value
        if descricao is None:
            continue

        print(f"busca item {cod} {descricao} na linha {x}")

        # busca inicial e final na composição
        linha_ini_comp = buscar_palavra_com_linha(
            sheet_comp,
            col_desc_comp,
            f"{cod} {descricao}",
            linha_busca_ini,
            sheet_comp_max,
        )
        if linha_ini_comp == -1:
            linha_ini_comp = buscar_palavra_com_linha(
                sheet_comp, col_desc_comp, cod, linha_busca_ini, sheet_comp_max
            )

        if linha_ini_comp == -1:
            linha_ini_comp = buscar_palavra_contem(
                sheet_comp, col_desc_comp, cod, linha_busca_ini, sheet_comp_max
            )

        if linha_ini_comp == -1:
            tk.messagebox.showwarning(
                "Aviso", f"Não foi encontrado o item na composição: {cod} {descricao}"
            )
            print(f"❌ Nao encontrado na composicao: {cod} {descricao}")
            continue

        print(
            f"encontrado item {cod} {descricao} -> linha da composicao: {linha_ini_comp}"
        )
        linha_fim_comp = buscar_palavra_com_linha(
            sheet_comp, col_totais_comp, valor_com_bdi, linha_ini_comp, sheet_comp_max
        )
        criar_link_composicao(
            sheet_origem=sheet,
            col_desc=col_desc,
            linha_origem=x,
            nome_planilha_destino=get_planilha_comp(dados),
            col_valor=col_valor_comp,
            linha_destino=linha_fim_comp,
        )
        linha_busca_ini = 1

        sheet[f"{col_preco_planilha}{x}"].value = (
            f"={get_planilha_comp(dados)}!{col_valor_comp}{linha_fim_comp}"
        )

        final_total_linha_array = set()
        for item in itens_array:
            res = fator_nos_item_totais(
                sheet_comp,
                dados,
                linha_ini_comp,
                linha_fim_comp,
                item["nome"],
                item["total"],
                item["fatorCoeficiente"] == "Sim",
                item["adicionarFator"] == "Sim",
                item.get("iniciaPor"),
                item.get("naoIniciaPor"),
            )
            if res:
                linha_desc, linha_total = res
                final_total_linha_array.add(linha_total)

                if item.get("buscarAuxiliar") == "Sim":
                    buscar_comp_auxiliar(
                        workbook, dados, itemChave, linha_desc, linha_total
                    )

        # soma final
        if final_total_linha_array:
            linha_valor_sum = buscar_palavra_com_linha(
                sheet_comp,
                col_totais_comp,
                valor_string,
                linha_ini_comp,
                linha_fim_comp,
            )
            if linha_valor_sum > 0:
                sheet_comp[f"{col_valor_comp}{linha_valor_sum}"].value = (
                    f"=SUM({','.join(f'{col_valor_comp}{linha}' for linha in final_total_linha_array)})"
                )
            else:
                print("linha_valor_sum <= 0")


def adicionar_fator_comp(workbook, dados, itemChave, lin_ini, lin_fim):
    sheet_comp = workbook[get_planilha_comp(dados)]

    copiar_colunas(sheet_comp, dados)
    adicionar_formula_preco_unitario_menos_preco_antigo(sheet_comp, dados)
    adicionar_fator_totais(workbook, dados, itemChave, lin_ini, lin_fim)
    valor_bdi_final(
        sheet_comp, dados, get_coluna_totais_comp(dados), get_valor_totais_comp(dados)
    )
