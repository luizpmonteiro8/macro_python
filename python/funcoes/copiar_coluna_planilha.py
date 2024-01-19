from funcoes.copiar_coluna import copiar_coluna


def copiar_coluna_planilha(sheet, dados):
    # Obter informações de coluna do JSON
    coluna_origem = dados.get(
        'colunaParaCopiar', {}).get('de', 'G')
    coluna_destino = dados.get(
        'colunaParaCopiar', {}).get('para', 'K')

    copiar_coluna(sheet, coluna_origem, coluna_destino)
