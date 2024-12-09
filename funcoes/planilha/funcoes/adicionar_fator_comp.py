import tkinter as tk

from openpyxl.utils import column_index_from_string

from funcoes.common.buscar_palavras import (buscar_palavra_com_linha,
                                            buscar_palavra_com_linha_exato)
from funcoes.common.copiar_coluna import copiar_coluna_com_numeros
from funcoes.common.valor_bdi_final import valor_bdi_final
from funcoes.get.get_linhas_json import (get_coeficiente_comp,
                                         get_coluna_totais_aux,
                                         get_coluna_totais_comp,
                                         get_copiar_coeficiente_comp,
                                         get_copiar_preco_unitario_comp,
                                         get_descricao_aux, get_descricao_comp,
                                         get_item_descricao_comp_aux,
                                         get_planilha_aux, get_planilha_codigo,
                                         get_planilha_comp,
                                         get_planilha_descricao,
                                         get_planilha_orcamentaria,
                                         get_planilha_preco_unitario,
                                         get_preco_unitario_comp,
                                         get_valor_com_bdi_string,
                                         get_valor_string,
                                         get_valor_totais_aux,
                                         get_valor_totais_comp)
from funcoes.planilha.funcoes.adicionar_fator_aux import \
    adicionar_fator_totais_aux


def copiar_colunas(sheet, dados):
    # Obter informações de coluna do JSON
    coluna_coeficiente = get_coeficiente_comp(dados)
    coluna_destino = get_copiar_coeficiente_comp(dados)

    coluna_preco_unitario = get_preco_unitario_comp(dados)
    coluna_destino1 = get_copiar_preco_unitario_comp(dados)

    copiar_coluna_com_numeros(sheet, coluna_coeficiente, coluna_destino)
    copiar_coluna_com_numeros(sheet, coluna_preco_unitario, coluna_destino1)


def adicionar_formula_preco_unitario_menos_preco_antigo(sheet, dados):
    coluna_origem = get_copiar_preco_unitario_comp(dados)
    coluna_preco_unitario = get_preco_unitario_comp(dados)
    linha_ini = 1
    final_linha = sheet.max_row + 1

    for x in range(linha_ini, final_linha):
        if sheet[f'{coluna_origem}{x}'].value is not None:
            coluna_destino = column_index_from_string(coluna_origem) + 1
            formula = f'=({coluna_origem}{x}-{coluna_preco_unitario}{x})'
            sheet.cell(row=x, column=coluna_destino).value = formula


def fator_nos_item_totais(sheet, dados,
                          linha_inicial_comp,
                          linha_final_comp,
                          nome,
                          totalNome,
                          coeficiente,
                          adicionar_fator):
    coluna_descricao_composicao = get_descricao_comp(dados)
    coluna_totais_composicao = get_coluna_totais_comp(dados)
    coluna_totais_valor_composicao = get_valor_totais_comp(dados)
    coluna_preco_unit = get_preco_unitario_comp(dados)
    coluna_coefieciente = get_coeficiente_comp(dados)
    coluna_preco_unitario_antigo = get_copiar_preco_unitario_comp(dados)
    coluna_coeficiente_antigo = get_copiar_coeficiente_comp(dados)

    # verifica se tem material na composicao
    inicial = buscar_palavra_com_linha_exato(
        sheet, coluna_descricao_composicao,
        nome,
        linha_inicial_comp, linha_final_comp)
    final = buscar_palavra_com_linha_exato(
        sheet, coluna_totais_composicao,
        totalNome,
        linha_inicial_comp, linha_final_comp)

    if (inicial > -1 and final > -1
        and inicial < linha_final_comp
            and inicial > linha_inicial_comp):
        # total final
        soma_formula = (
            f'=SUM('
            f'{coluna_totais_valor_composicao}{inicial+1}:'
            f'{coluna_totais_valor_composicao}{final-1}'
            f')'
        )
        sheet[f'{coluna_totais_valor_composicao}{final}'].value = soma_formula
        for y in range(inicial+1, final):
            # total linha
            #sheet[f'{coluna_totais_valor_composicao}{y}'].value = (
            #    f'=ROUND({coluna_coefieciente}{y}*{coluna_preco_unit}{y}, 2)'
            #)
            if coeficiente and adicionar_fator:
                sheet[f'{coluna_coefieciente}{y}'].value = (
                    f'={coluna_coeficiente_antigo}{y}*FATOR'
                )
            else:
                if adicionar_fator:
                    sheet[f'{coluna_preco_unit}{y}'].value = (
                        f'=ROUND({coluna_preco_unitario_antigo}{y}*FATOR, 2)'
                    )

        return inicial, final


def buscar_comp_auxiliar(workbook, dados, itemChave, linha, linha_total):
    sheet_name_comp = get_planilha_comp(dados)
    sheet_planilha_comp = workbook[sheet_name_comp]
    sheet_name_aux = get_planilha_aux(dados)
    sheet_planilha_aux = workbook[sheet_name_aux]

    coluna_item = get_item_descricao_comp_aux(dados)
    coluna_desc_aux = get_descricao_aux(dados)
    coluna_totais_aux = get_coluna_totais_aux(dados)
    coluna_valor_aux = get_valor_totais_aux(dados)
    coluna_preco_unit = get_preco_unitario_comp(dados)
    valor_string = get_valor_string(dados)

    ultima_linha_busca = 1
    ultima_linha_aux = sheet_planilha_aux.max_row

    # noma para busca na auxiliar
    for x in range(linha, linha_total):
        cod = sheet_planilha_comp[f'{coluna_desc_aux}{x}'].value
        nome = sheet_planilha_comp[f'{coluna_item}{x}'].value

        if nome is not None:
            linha_inicial = buscar_palavra_com_linha(
                sheet_planilha_aux, coluna_desc_aux, cod + ' ' + nome,
                ultima_linha_busca, ultima_linha_aux)

            if (linha_inicial == -1):
                linha_inicial = buscar_palavra_com_linha(
                    sheet_planilha_aux, coluna_desc_aux, cod,
                    ultima_linha_busca, ultima_linha_aux)

            if linha_inicial > -1:
                linha_final = buscar_palavra_com_linha(
                    sheet_planilha_aux, coluna_totais_aux, valor_string,
                    linha_inicial, ultima_linha_aux
                )

                # adiciona fator e totais no auxiliar
                adicionar_fator_totais_aux(
                    workbook, dados, itemChave, linha_inicial, linha_final
                )
                # coloca valor do auxiliar na composicao
                sheet_planilha_comp[f'{coluna_preco_unit}{x}'].value = (
                    f'=\'{sheet_name_aux}\'!{coluna_valor_aux}{linha_final}')


def adicionar_fator_totais(workbook, dados, itemChave, linhaIni, linhaFim):
    sheet_name = get_planilha_orcamentaria(dados)
    sheet_planilha = workbook[sheet_name]
    sheet_name_comp = get_planilha_comp(dados)
    sheet_planilha_comp = workbook[sheet_name_comp]
    sheet_comp_linha_fim = sheet_planilha_comp.max_row + 1

    coluna_preco_planilha = get_planilha_preco_unitario(dados)
    coluna_descricao_composicao = get_descricao_comp(dados)
    coluna_totais_composicao = get_coluna_totais_comp(dados)

    valor_com_bdi = get_valor_com_bdi_string(dados)
    valorString = get_valor_string(dados)
    coluna_cod = get_planilha_codigo(dados)
    coluna_descricao = get_planilha_descricao(dados)
    coluna_totais_comp = get_coluna_totais_comp(dados)
    coluna_valor_string = get_valor_totais_comp(dados)

    itens_array = []

    # Iterar sobre as chaves que começam com "item"
    for chave, valor in itemChave.items():
        if chave.startswith("item"):
            itens_array.append(valor)

    # evitar usar valor errado iniciando no ultimo que foi buscado
    linha_final_iniciar_busca = 1

    for x in range(linhaIni, linhaFim):
        # busca nome da descricao na planilha orcamentaria
        cod = sheet_planilha[f'{coluna_cod}{x}'].value
        coluna_busca_value = sheet_planilha[f'{coluna_descricao}{x}'].value

        if coluna_busca_value is not None:
            # busca nome da descricao na composicao
            linha_inicial_comp = -1
            linha_inicial_comp = buscar_palavra_com_linha(
                sheet_planilha_comp, coluna_descricao_composicao,
                cod + ' ' + coluna_busca_value, linha_final_iniciar_busca,
                sheet_comp_linha_fim)
            if (linha_inicial_comp == -1):
                linha_inicial_comp = buscar_palavra_com_linha(
                    sheet_planilha_comp, coluna_descricao_composicao,
                    cod, linha_final_iniciar_busca,
                    sheet_comp_linha_fim)

            if (linha_inicial_comp == -1):
                tk.messagebox.showwarning(
                    "Aviso",
                    "Não foi encontrado o item na composição. " +
                    cod + " " + coluna_busca_value +
                    "Verifique se o item existe em composição. ")

            # busca linha final pelo valor bdi
            linha_final_comp = buscar_palavra_com_linha(
                sheet_planilha_comp, coluna_totais_comp, valor_com_bdi,
                linha_inicial_comp, sheet_comp_linha_fim)
            # linha por onde ele vai buscar inicialmente na proxima rodada
            linha_final_iniciar_busca = linha_final_comp

            # adicionando formula no preco unitario em planilha
            sheet_planilha[f'{coluna_preco_planilha}{x}'].value = (
                f'={sheet_name_comp}!{coluna_valor_string}{linha_final_comp}')

            final_total_linha_array = []

            for item in itens_array:
                resultado_fator = fator_nos_item_totais(
                    sheet_planilha_comp, dados,
                    linha_inicial_comp,
                    linha_final_comp,
                    item['nome'],
                    item['total'],
                    True if item['fatorCoeficiente'] == 'Sim' else False,
                    True if item['adicionarFator'] == 'Sim' else False,
                )

                if resultado_fator is not None:
                    linha_desc, linha_total = resultado_fator
                if resultado_fator is not None and linha_total is not None:
                    final_total_linha_array.append(linha_total)

                # busca auxiliar
                if (item['buscarAuxiliar'] is not None
                        and item['buscarAuxiliar'] == 'Sim'
                        and resultado_fator is not None
                        and linha_desc > 0
                        and linha_total > 0):
                    buscar_comp_auxiliar(
                        workbook, dados, itemChave, linha_desc, linha_total)

            # total no VALOR:
            if final_total_linha_array:
                linha_valor_sum = buscar_palavra_com_linha(
                    sheet_planilha_comp, coluna_totais_composicao, valorString,
                    linha_inicial_comp, linha_final_comp)

                if linha_valor_sum > 0:
                    formula_soma = (
                        '=SUM(' +
                        ','.join([f'{coluna_valor_string}{linha}'
                                  for linha in final_total_linha_array]) +
                        ')'
                    )

                    # Atribui a fórmula à célula específica
                    sheet_planilha_comp[
                        f'{coluna_valor_string}{linha_valor_sum}'
                    ].value = formula_soma
                else:
                    print("A linha_valor_sum não é maior que zero.")


def adicionar_fator_comp(workbook, dados, itemChave, linhaIniPlan,
                         linhaFimPlan):

    sheet_name_comp = get_planilha_comp(dados)
    sheet_planilha_comp = workbook[sheet_name_comp]

    copiar_colunas(sheet_planilha_comp, dados)
    adicionar_formula_preco_unitario_menos_preco_antigo(
        sheet_planilha_comp, dados)
    adicionar_fator_totais(workbook, dados, itemChave,
                           linhaIniPlan, linhaFimPlan)
    coluna_valor_string = get_coluna_totais_comp(dados)
    coluna_valor_value = get_valor_totais_comp(dados)
    valor_bdi_final(sheet_planilha_comp, dados,
                    coluna_valor_string, coluna_valor_value)
