def atualizar_linha_total(sheet, linha_valor):
    valor_h = sheet[f'H{linha_valor}'].value
    sheet[f'K{linha_valor}'].value = valor_h
    sheet[f'L{linha_valor}'].value = f'=H{linha_valor}-K{linha_valor}'


def somatario_int(sheet, ultima_linha):
    # Índice fixo da coluna
    indice_coluna = 1  # 1 para a coluna 'A'

    # Lista para armazenar valores que começam com o inteiro selecionado
    valores_selecionados = []
    linha_valor_int = 0
    linha_valor_int_todas = []

    valores_selecionados_1 = []
    linha_valor_int_1 = 0
    valor_inicial = ""

    # Iterar sobre as linhas da planilha
    for linha in sheet.iter_rows(min_row=1, max_row=ultima_linha):

        # Obter o valor da célula na coluna 'A'
        valor = linha[indice_coluna - 1].value

        # Verificar se o valor é um número inteiro e
        # menor ou igual ao maior_int
        if (valor is not None and str(valor).count('.') == 0
                and valor != 'ITEM'):
            # Salvar o valor inteiro em uma variável
            if linha_valor_int == 0:
                linha_valor_int = linha[0].row
                linha_valor_int_todas.append(f"H{linha[0].row}")
            else:
                formula_somatorio = f'=SUM({",".join(valores_selecionados)})'
                atualizar_linha_total(sheet, linha_valor_int)
                sheet[f'H{linha_valor_int}'].value = formula_somatorio
                linha_valor_int = linha[0].row
                linha_valor_int_todas.append(f"H{linha[0].row}")
                valores_selecionados = []

        # Verificar se o valor começa com o inteiro selecionado
        # e tem apenas um ponto decimal
        if (str(valor).count('.') == 1):
            valores_selecionados.append(f"H{linha[0].row}")
            if linha_valor_int_1 != 0:
                formula_somatorio = f'=SUM({",".join(valores_selecionados_1)})'
                atualizar_linha_total(sheet, linha_valor_int_1-1)
                sheet[f'H{linha_valor_int_1-1}'].value = formula_somatorio
                linha_valor_int_1 = 0
                valores_selecionados_1 = []
                valor_inicial = ''
                continue

        if str(valor).count('.') == 2:
            if valor_inicial == '':
                valor_corte = str(valor).split('.')
                valor_inicial = valor_corte[0] + '.' + valor_corte[1]
                valores_selecionados_1.append(f"H{linha[0].row}")
                if linha_valor_int_1 == 0:
                    linha_valor_int_1 = linha[0].row
                continue

            if str(valor).startswith(valor_inicial):
                valores_selecionados_1.append(f"H{linha[0].row}")
                continue

        if valor is None and linha[0].row == ultima_linha - 2:
            linha_valor_bdi = ultima_linha-2
            linha_valor_orcamento = ultima_linha - 1
            linha_valor_total = ultima_linha

            atualizar_linha_total(sheet, linha_valor_bdi)
            atualizar_linha_total(sheet, linha_valor_orcamento)
            atualizar_linha_total(sheet, linha_valor_total)

            sheet[f'H{linha_valor_orcamento}'].value = (f'=SUM({
                ",".join(linha_valor_int_todas)})')
            sheet[f'H{linha_valor_bdi}'].value = (
                f'=ROUND(H{linha_valor_orcamento}*BDI,2)')
            sheet[f'H{linha_valor_total}'].value = (
                f'=H{linha_valor_orcamento}+H{linha_valor_bdi}')

            # escreve ultimo
            formula_somatorio = f'=SUM({",".join(valores_selecionados)})'
            atualizar_linha_total(sheet, linha_valor_int)
            sheet[f'H{linha_valor_int}'].value = formula_somatorio
            linha_valor_int = linha[0].row
            linha_valor_int_todas.append(f"H{linha[0].row}")
            valores_selecionados = []


def somatorio_planilha(sheet):
    ultima_linha = sheet.max_row

    somatario_int(sheet, ultima_linha)
