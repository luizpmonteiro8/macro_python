def copiar_coluna(sheet, coluna_origem, coluna_destino):
    # Obtém os valores da coluna de origem
    valores_coluna = [sheet[f'{coluna_origem}{linha}'].value
                      for linha in range(1, sheet.max_row + 1)]

    # Copia os valores não nulos para a coluna de destino
    for linha, valor in enumerate(valores_coluna, start=1):
        if valor is not None:
            sheet[f'{coluna_destino}{linha}'] = valor


def copiar_coluna_com_numeros(sheet, coluna_origem, coluna_destino):
    # Obtém os valores da coluna de origem
    valores_coluna = [sheet[f'{coluna_origem}{linha}'].value
                      for linha in range(1, sheet.max_row + 1)]

    # Copia os valores numéricos para a coluna de destino
    for linha, valor in enumerate(valores_coluna, start=1):
        if valor is not None and str(valor).replace('.', '', 1).isdigit():
            sheet[f'{coluna_destino}{linha}'] = valor
