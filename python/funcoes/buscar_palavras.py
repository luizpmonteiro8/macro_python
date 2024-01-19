def buscar_palavra(sheet, coluna, palavra):
    # Converte a letra da coluna para o número correspondente (A=1, B=2, etc.)
    numero_coluna = ord(coluna.upper()) - ord('A') + 1

    # Percorre as células da coluna e verifica se a palavra está presente
    for linha in range(1, sheet.max_row + 1):
        valor_celula = sheet.cell(row=linha, column=numero_coluna).value
        if valor_celula is not None:
            if palavra.lower() in str(valor_celula).lower():
                return linha

    # Se a palavra não foi encontrada, retorna algum valor indicativo, como -1
    return -1


def buscar_palavra_com_linha(sheet, coluna, palavra, linha_inicial,
                             linha_final):
    # Converte a letra da coluna para o número correspondente (A=1, B=2, etc.)
    numero_coluna = ord(coluna.upper()) - ord('A') + 1

    # Define a linha final como a última linha se não for fornecida
    linha_final = linha_final or sheet.max_row

    # Percorre as células da coluna e verifica se a palavra está presente
    for linha in range(linha_inicial, linha_final):
        valor_celula = sheet.cell(row=linha, column=numero_coluna).value
        if valor_celula is not None:
            if palavra.lower() in str(valor_celula).lower():
                return linha

    # Se a palavra não foi encontrada, retorna algum valor indicativo, como -1
    return -1
