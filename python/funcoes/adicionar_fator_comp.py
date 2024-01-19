from openpyxl.utils import column_index_from_string

from funcoes.buscar_palavras import buscar_palavra, buscar_palavra_com_linha
from funcoes.copiar_coluna import copiar_coluna_com_numeros


def copiar_colunas(sheet, dados):
    # Obter informações de coluna do JSON
    coluna_origem = dados.get(
        'colunaParaCopiarComposicao', {}).get('de', 'E')
    coluna_destino = dados.get(
        'colunaParaCopiarComposicao', {}).get('para', 'L')

    coluna_origem1 = dados.get(
        'colunaParaCopiarComposicao1', {}).get('de', 'F')
    coluna_destino1 = dados.get(
        'colunaParaCopiarComposicao1', {}).get('para', 'M')

    copiar_coluna_com_numeros(sheet, coluna_origem, coluna_destino)
    copiar_coluna_com_numeros(sheet, coluna_origem1, coluna_destino1)


def adicionar_formula_preco_unitario_menos_preco_antigo(sheet, dados):
    coluna_origem = dados.get(
        'colunaParaCopiarComposicao1', {}).get('para', 'M')
    coluna_preco_unitario = dados.get(
        'composicaoPrecoUnitario', 'F')
    linha_ini = 1
    final_linha = sheet.max_row + 1

    for x in range(linha_ini, final_linha):
        if sheet[f'{coluna_origem}{x}'].value is not None:
            coluna_destino = column_index_from_string(coluna_origem) + 1
            formula = f'=({coluna_origem}{x}-{coluna_preco_unitario}{x})'
            sheet.cell(row=x, column=coluna_destino).value = formula


def fator_nos_item(sheet, dados, linha_inicial_comp, linha_final_comp, nome,
                   totalNome, coeficiente):
    coluna_descricao_composicao = dados.get(
        'colunaDescricaoComposicao', 'A'
    )
    coluna_totais_composicao = dados.get(
        'colunaTotaisComposicao', 'E'
    )
    coluna_preco_unit = dados.get(
        "composicaoPrecoUnitario", "F"
    )
    coluna_coefieciente = dados.get(
        "composicaoCoeficiente", "E"
    )
    coluna_preco_unitario_antigo = dados.get(
        'colunaParaCopiarComposicao1', {}).get('para', 'M')
    coluna_coeficiente_antigo = dados.get(
        'colunaParaCopiarComposicao', {}).get('para', 'L')

    # verifica se tem material na composicao
    inicial = buscar_palavra_com_linha(
        sheet, coluna_descricao_composicao,
        nome,
        linha_inicial_comp, linha_final_comp)
    final = buscar_palavra_com_linha(
        sheet, coluna_totais_composicao,
        totalNome,
        linha_inicial_comp, linha_final_comp)

    if (inicial > -1 and final > -1
        and inicial < linha_final_comp
            and inicial > linha_inicial_comp):
        for y in range(inicial+1, final):
            if coeficiente:
                sheet[f'{coluna_coefieciente}{y}'].value = (
                    f'={coluna_coeficiente_antigo}{y}*FATOR'
                )
            else:
                sheet[f'{coluna_preco_unit}{y}'].value = (
                    f'=ROUND({coluna_preco_unitario_antigo}{y}*FATOR, 2)'
                )


def adicionar_fator(workbook, dados, linhaIni, linhaFim):
    sheet_name = dados.get(
        'planilha', 'PLANILHA ORCAMENTARIA')
    sheet_planilha = workbook[sheet_name]
    sheet_name_comp = dados.get(
        'planilhaComposicao', 'COMPOSICAO')
    sheet_planilha_comp = workbook[sheet_name_comp]
    sheet_comp_linha_fim = sheet_planilha_comp.max_row

    coluna_descricao_composicao = dados.get(
        'colunaDescricaoComposicao', 'A'
    )

    # verificar
    valor_com_bdi = dados.get(
        'valorComBdi', 'VALOR COM BDI:')
    coluna_descricao = dados.get(
        'colunaDescricaoEmPlanilhaParaBuscaEmComposicao', 'C'
    )
    coluna_valor_bdi = dados.get(
        'colunaValorComBdi', 'E'
    )

    itens_array = []

    # Iterar sobre as chaves que começam com "item"
    for chave, valor in dados.items():
        if chave.startswith("item"):
            itens_array.append(valor)

    for x in range(linhaIni, linhaFim):
        # busca nome da descricao na planilha
        coluna_busca_value = sheet_planilha[f'{coluna_descricao}{x}'].value

        if coluna_busca_value is not None:
            # busca nome da descricao na composicao
            linha_inicial_comp = buscar_palavra(
                sheet_planilha_comp, coluna_descricao_composicao,
                coluna_busca_value)
            # busca linha final pelo valor bdi
            linha_final_comp = buscar_palavra_com_linha(
                sheet_planilha_comp, coluna_valor_bdi, valor_com_bdi,
                linha_inicial_comp, sheet_comp_linha_fim)

            for item in itens_array:
                fator_nos_item(sheet_planilha_comp, dados, linha_inicial_comp,
                               linha_final_comp,
                               item['nome'], item['total'],
                               True if item['fatorCoeficiente'] == 'Sim' else
                               False)


def adicionar_fator_comp(workbook, dados, linhaIni, linhaFim):
    sheet_name_comp = dados.get(
        'planilhaComposicao', 'COMPOSICAO')
    sheet_planilha_comp = workbook[sheet_name_comp]

    copiar_colunas(sheet_planilha_comp, dados)
    adicionar_formula_preco_unitario_menos_preco_antigo(
        sheet_planilha_comp, dados)
    adicionar_fator(workbook, dados, linhaIni, linhaFim)
