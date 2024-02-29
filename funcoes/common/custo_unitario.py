from funcoes.common.buscar_palavras import buscar_palavra_com_linha_exato
from funcoes.get.get_linhas_json import (get_coluna_totais_aux,
                                         get_coluna_totais_comp,
                                         get_planilha_aux, get_planilha_comp,
                                         get_valor_totais_aux,
                                         get_valor_totais_comp)

# buscar Custo Horário da Execução: na planilha comp e aux


def custo_unitario_execucao(workbook, dados):
    sheet_name_comp = get_planilha_comp(dados)
    sheet_planilha_comp = workbook[sheet_name_comp]
    sheet_name_aux = get_planilha_aux(dados)
    sheet_planilha_aux = workbook[sheet_name_aux]
    coluna_totais_comp = get_coluna_totais_comp(dados)
    coluna_totais_aux = get_coluna_totais_aux(dados)
    col_valor_comp = get_valor_totais_comp(dados)
    coluna_valor_aux = get_valor_totais_aux(dados)

    linha_final_comp = sheet_planilha_comp.max_row
    linha_final_aux = sheet_planilha_aux.max_row

    custo_horario_composicao = []
    custo_horario_aux = []

    linha_vazia = 0
    for x in range(1, linha_final_comp):
        valor = sheet_planilha_comp[f'{coluna_totais_comp}{x}'].value
        if (valor == '' or valor is None):
            linha_vazia = x

        if (valor ==
                'Custo Horário da Execução:'):
            custo_horario_composicao.append(
                {'linha_vazia': linha_vazia, 'linha_custo': x})

    for x in range(1, linha_final_aux):
        valor = sheet_planilha_aux[f'{coluna_totais_aux}{x}'].value

        if (valor == '' or valor is None):
            linha_vazia = x
        if (valor == 'Custo Horário da Execução:'):
            custo_horario_aux.append(
                {'linha_vazia': linha_vazia, 'linha_custo': x})

    if len(custo_horario_composicao) > 0:
        for x in custo_horario_composicao:
            linha_inicial = x['linha_vazia']
            linha_custo = x['linha_custo']

            string1 = 'TOTAL MÃO DE OBRA:'
            string2 = 'TOTAL EQUIPAMENTOS:'

            linha_total_mao = buscar_palavra_com_linha_exato(
                sheet_planilha_comp,
                coluna_totais_comp,
                string1, linha_inicial,
                linha_custo
            )

            linha_total_equipamento = buscar_palavra_com_linha_exato(
                sheet_planilha_comp,
                coluna_totais_comp,
                string2,
                linha_inicial,
                linha_custo
            )

            if linha_total_mao != -1 and linha_total_equipamento != -1:
                formula_soma = f'={col_valor_comp}{linha_total_mao}+' \
                    f'{col_valor_comp}{linha_total_equipamento}'
            elif linha_total_mao != -1:
                formula_soma = f'={col_valor_comp}{linha_total_mao}'
            elif linha_total_equipamento != -1:
                formula_soma = f'={col_valor_comp}{
                    linha_total_equipamento}'

            sheet_planilha_comp[f'{col_valor_comp}{linha_custo}'].value = (
                formula_soma)
            sheet_planilha_comp[f'{col_valor_comp}{linha_custo+2}'].value = (
                f'=ROUND({col_valor_comp}' +
                f'{linha_custo}/{col_valor_comp}{linha_custo+1}, 4)'
            )

    if len(custo_horario_aux) > 0:
        for x in custo_horario_aux:
            linha_inicial = x['linha_vazia']
            linha_custo = x['linha_custo']

            string1 = 'TOTAL MÃO DE OBRA:'
            string2 = 'TOTAL EQUIPAMENTOS:'

            linha_total_mao = buscar_palavra_com_linha_exato(
                sheet_planilha_aux,
                coluna_totais_aux,
                string1, linha_inicial,
                linha_custo
            )

            linha_total_equipamento = buscar_palavra_com_linha_exato(
                sheet_planilha_aux,
                coluna_totais_aux,
                string2,
                linha_inicial,
                linha_custo
            )

            if linha_total_mao != -1 and linha_total_equipamento != -1:
                formula_soma = f'={coluna_valor_aux}{linha_total_mao}+' \
                    f'{coluna_valor_aux}{linha_total_equipamento}'
            elif linha_total_mao != -1:
                formula_soma = f'={coluna_valor_aux}{linha_total_mao}'
            elif linha_total_equipamento != -1:
                formula_soma = f'={coluna_valor_aux}{
                    linha_total_equipamento}'

            sheet_planilha_aux[f'{coluna_valor_aux}{linha_custo}'].value = (
                formula_soma)
            sheet_planilha_aux[f'{coluna_valor_aux}{linha_custo+2}'].value = (
                f'=ROUND({coluna_valor_aux}' +
                f'{linha_custo}/{coluna_valor_aux}{linha_custo+1}, 4)'
            )
