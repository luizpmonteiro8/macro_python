from openpyxl.utils import column_index_from_string

from funcoes.common.buscar_palavras import (
    buscar_palavra_com_linha,
    buscar_palavra_com_linha_exato,
)
from funcoes.common.copiar_coluna import copiar_coluna_com_numeros
from funcoes.common.valor_bdi_final import valor_bdi_final
from funcoes.get.get_linhas_json import (
    get_coeficiente_aux,
    get_coluna_totais_aux,
    get_copiar_coeficiente_aux,
    get_copiar_preco_unitario_aux,
    get_descricao_aux,
    get_item_descricao_comp_aux,
    get_planilha_aux,
    get_preco_unitario_aux,
    get_valor_string,
    get_valor_totais_aux,
)


def copiar_colunas(sheet, dados):
    # Obter informações de coluna do JSON
    coluna_origem = get_coeficiente_aux(dados)
    coluna_destino = get_copiar_coeficiente_aux(dados)

    coluna_origem1 = get_preco_unitario_aux(dados)
    coluna_destino1 = get_copiar_preco_unitario_aux(dados)

    copiar_coluna_com_numeros(sheet, coluna_origem, coluna_destino)
    copiar_coluna_com_numeros(sheet, coluna_origem1, coluna_destino1)


def adicionar_formula_preco_unitario_menos_preco_antigo(sheet, dados):
    coluna_origem = get_copiar_preco_unitario_aux(dados)
    coluna_preco_unitario_aux = get_preco_unitario_aux(dados)
    linha_ini = 1
    final_linha = sheet.max_row + 1

    for x in range(linha_ini, final_linha):
        if sheet[f"{coluna_origem}{x}"].value is not None:
            coluna_destino = column_index_from_string(coluna_origem) + 1
            formula = f"=({coluna_origem}{x}-{coluna_preco_unitario_aux}{x})"
            sheet.cell(row=x, column=coluna_destino).value = formula


def fator_nos_item_totais_aux(
    sheet,
    dados,
    linha_inicial_comp,
    linha_final_comp,
    nome,
    totalNome,
    coeficiente,
    adicionar_fator,
    inicia_por=None,
    nao_inicia_por=None,
):
    coluna_descricao_aux = get_descricao_aux(dados)
    coluna_totais_aux = get_coluna_totais_aux(dados)
    coluna_totais_valor_aux = get_valor_totais_aux(dados)
    coluna_preco_unit = get_preco_unitario_aux(dados)
    coluna_coefieciente = get_coeficiente_aux(dados)
    coluna_preco_unitario_antigo = get_copiar_preco_unitario_aux(dados)
    coluna_coeficiente_antigo = get_copiar_coeficiente_aux(dados)

    # verifica se tem material na composicao
    inicial = buscar_palavra_com_linha_exato(
        sheet, coluna_descricao_aux, nome, linha_inicial_comp, linha_final_comp
    )
    final = buscar_palavra_com_linha_exato(
        sheet, coluna_totais_aux, totalNome, linha_inicial_comp, linha_final_comp
    )

    if (
        inicial > -1
        and final > -1
        and inicial < linha_final_comp
        and inicial > linha_inicial_comp
    ):
        # total final
        soma_formula = (
            f"=SUM("
            f"{coluna_totais_valor_aux}{inicial+1}:"
            f"{coluna_totais_valor_aux}{final-1}"
            f")"
        )
        sheet[f"{coluna_totais_valor_aux}{final}"].value = soma_formula
        for y in range(inicial + 1, final):
            if inicia_por:
                descricao_atual = sheet[f"{coluna_descricao_aux}{y}"].value
                if descricao_atual is None or not descricao_atual.startswith(
                    inicia_por
                ):
                    continue
            if nao_inicia_por:
                descricao_atual = sheet[f"{coluna_descricao_aux}{y}"].value
                if descricao_atual is None or descricao_atual.startswith(
                    nao_inicia_por
                ):
                    continue
            if coeficiente and adicionar_fator:
                sheet[f"{coluna_coefieciente}{y}"].value = (
                    f"={coluna_coeficiente_antigo}{y}*FATOR"
                )
            else:
                if adicionar_fator:
                    sheet[f"{coluna_preco_unit}{y}"].value = (
                        f"=ROUND({coluna_preco_unitario_antigo}{y}*FATOR, 2)"
                    )

        return inicial, final


def buscar_auxiliar_no_aux(workbook, dados, itemChave, linha, linha_total):
    # busca dentro de auxiliar os auxiliares
    sheet_name_aux = get_planilha_aux(dados)
    sheet_planilha_aux = workbook[sheet_name_aux]

    coluna_item = get_item_descricao_comp_aux(dados)
    coluna_desc_aux = get_descricao_aux(dados)
    coluna_valor_aux = get_valor_totais_aux(dados)
    coluna_preco_aux = get_preco_unitario_aux(dados)
    coluna_totais_aux = get_coluna_totais_aux(dados)
    valor_string = get_valor_string(dados)

    coluna_preco_unitario_antigo = get_copiar_preco_unitario_aux(dados)

    ultima_linha = sheet_planilha_aux.max_row

    itens_array = []

    # Iterar sobre as chaves que começam com "item"
    for chave, valor in itemChave.items():
        if chave.startswith("item"):
            itens_array.append(valor)

    for x in range(linha, linha_total):
        cod = sheet_planilha_aux[f"{coluna_desc_aux}{x}"].value
        item = sheet_planilha_aux[f"{coluna_item}{x}"].value

        if item is not None:
            linha_inicial = buscar_palavra_com_linha(
                sheet_planilha_aux, coluna_desc_aux, cod + " " + item, 1, ultima_linha
            )

            if linha_inicial == -1:
                linha_inicial = buscar_palavra_com_linha(
                    sheet_planilha_aux, coluna_desc_aux, cod, 1, ultima_linha
                )

            if linha_inicial > -1:
                linha_final = buscar_palavra_com_linha_exato(
                    sheet_planilha_aux,
                    coluna_totais_aux,
                    valor_string,
                    linha_inicial,
                    ultima_linha,
                )

                if cod.startswith("I"):
                    # adicionando formula no preco unitario em auxiliar
                    sheet_planilha_aux[f"{coluna_preco_aux}{x}"].value = (
                        f"=ROUND({coluna_preco_unitario_antigo}{x}*FATOR, 2)"
                    )
                else:
                    # adicionando formula no preco unitario em auxiliar
                    sheet_planilha_aux[f"{coluna_preco_aux}{x}"].value = (
                        f"='{sheet_name_aux}'!{
                            coluna_valor_aux}{linha_final}"
                    )

                final_total_linha_array = []

                for item in itens_array:
                    resultado_fator = fator_nos_item_totais_aux(
                        sheet_planilha_aux,
                        dados,
                        linha_inicial,
                        linha_final,
                        item["nome"],
                        item["total"],
                        True if item["fatorCoeficiente"] == "Sim" else False,
                        True if item["adicionarFator"] == "Sim" else False,
                        item["iniciaPor"],
                        item["naoIniciaPor"],
                    )
                    if resultado_fator is not None:
                        linha_desc, linha_total = resultado_fator
                    if resultado_fator is not None and linha_total is not None:
                        final_total_linha_array.append(linha_total)

                    if (
                        item["buscarAuxiliar"] is not None
                        and item["buscarAuxiliar"] == "Sim"
                        and resultado_fator is not None
                        and linha_desc > 0
                        and linha_total > 0
                    ):
                        buscar_auxiliar_no_aux(
                            workbook, dados, itemChave, linha_desc, linha_total
                        )

                # total no VALOR:
                if final_total_linha_array:
                    linha_valor_sum = buscar_palavra_com_linha(
                        sheet_planilha_aux,
                        coluna_totais_aux,
                        valor_string,
                        linha_inicial,
                        linha_final + 1,
                    )

                    if linha_valor_sum > 0:
                        formula_soma = (
                            "=SUM("
                            + ",".join(
                                [
                                    f"{coluna_valor_aux}{linha}"
                                    for linha in final_total_linha_array
                                ]
                            )
                            + ")"
                        )

                        # Atribui a fórmula à célula específica
                        sheet_planilha_aux[
                            f"{coluna_valor_aux}{linha_valor_sum}"
                        ].value = formula_soma
                    else:
                        print("A linha_valor_sum não é maior que zero.")


def adicionar_fator_totais_aux(workbook, dados, itemChave, linhaIni, linhaFim):
    # chamado no adicionar_fator_comp
    sheet_name_aux = get_planilha_aux(dados)
    sheet_planilha_aux = workbook[sheet_name_aux]

    coluna_totais_aux = get_coluna_totais_aux(dados)

    valorString = get_valor_string(dados)
    coluna_valor_string = get_valor_totais_aux(dados)

    itens_array = []

    # Iterar sobre as chaves que começam com "item"
    for chave, valor in itemChave.items():
        if chave.startswith("item"):
            itens_array.append(valor)

    final_total_linha_array = []

    for item in itens_array:
        resultado_fator = fator_nos_item_totais_aux(
            sheet_planilha_aux,
            dados,
            linhaIni,
            linhaFim,
            item["nome"],
            item["total"],
            True if item["fatorCoeficiente"] == "Sim" else False,
            True if item["adicionarFator"] == "Sim" else False,
            item["iniciaPor"],
            item["naoIniciaPor"],
        )

        if resultado_fator is not None:
            linha_desc, linha_total = resultado_fator
        if resultado_fator is not None and linha_total is not None:
            final_total_linha_array.append(linha_total)
        if (
            item["buscarAuxiliar"] is not None
            and item["buscarAuxiliar"] == "Sim"
            and resultado_fator is not None
            and linha_desc > 0
            and linha_total > 0
        ):
            buscar_auxiliar_no_aux(workbook, dados, itemChave, linha_desc, linha_total)

        # total no VALOR:
        if final_total_linha_array:
            linha_valor_sum = buscar_palavra_com_linha(
                sheet_planilha_aux,
                coluna_totais_aux,
                valorString,
                linhaIni,
                linhaFim + 1,
            )
            if linha_valor_sum > 0:
                formula_soma = (
                    "=SUM("
                    + ",".join(
                        [
                            f"{coluna_valor_string}{linha}"
                            for linha in final_total_linha_array
                        ]
                    )
                    + ")"
                )
                # Atribui a fórmula à célula específica
                sheet_planilha_aux[f"{coluna_valor_string}{linha_valor_sum}"].value = (
                    formula_soma
                )
            else:
                print("A linha_valor_sum não é maior que zero.")


def adicionar_fator_aux(workbook, dados):
    sheet_name_aux = get_planilha_aux(dados)
    sheet_planilha_aux = workbook[sheet_name_aux]

    copiar_colunas(sheet_planilha_aux, dados)
    adicionar_formula_preco_unitario_menos_preco_antigo(sheet_planilha_aux, dados)
    coluna_valor_string = get_coluna_totais_aux(dados)
    coluna_valor_value = get_valor_totais_aux(dados)
    valor_bdi_final(sheet_planilha_aux, dados, coluna_valor_string, coluna_valor_value)
