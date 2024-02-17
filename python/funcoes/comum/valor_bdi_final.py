def valor_bdi_final(sheet, dados, coluna_valor_string, coluna_valor_value):
    valor_Bdi = dados.get('valorBdi', "VALOR BDI")
    linha_final = sheet.max_row

    for x in range(2, linha_final + 1):  # Começamos do segundo
        # para evitar índices negativos
        valor = sheet[f'{coluna_valor_string}{x}'].value

        if valor_Bdi in str(valor):
            # Adiciona a fórmula na linha atual
            formula_atual = (
                f'=ROUNDDOWN({coluna_valor_value}{x-1}*BDI,2)')
            sheet[f'{coluna_valor_value}{x}'].value = formula_atual

            # Adiciona a fórmula na próxima linha
            formula_proxima = f'={coluna_valor_value}{
                x} + {coluna_valor_value}{x-1}'
            sheet[f'{coluna_valor_value}{x + 1}'].value = formula_proxima
