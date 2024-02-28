from openpyxl.utils import column_index_from_string


# orcamentaria
def get_planilha_orcamentaria(dados):
    return dados.get('planilhaOrcamentaria', 'PLANILHA ORCAMENTARIA')


def get_coluna_inicial(dados):
    return dados.get('colunaInicial', 'A')


def get_coluna_final(dados):
    return dados.get('colunaFinal', 'F')


def get_valor_inicial(dados):
    return dados.get('valorInicial', 'ITEM')


def get_valor_final(dados):
    return dados.get('valorFinal', 'VALOR BDI TOTAL')


def get_planilha_quantidade(dados):
    return dados.get('planilhaQuantidade', 'F')


def get_planilha_preco_unitario(dados):
    return dados.get('planilhaPrecoUnitario', 'G')


def get_planilha_preco_total(dados):
    return dados.get('planilhaPrecoTotal', 'H')


def get_planilha_codigo(dados):
    return dados.get('planilhaCodigo', 'B')


def get_planilha_descricao(dados):
    return dados.get('planilhaDescricao', 'C')


def get_planilha_preco_unitario_copiar(dados):
    return dados.get('planilhaPrecoUnitarioCopiar', 'K')


# fator
def get_planilha_fator(dados):
    return dados.get('planilhaFator', 'RESUMO')


def get_valor_bdi(dados):
    return dados.get('BDI', '28.82')


def get_linha_fator(dados):
    return int(dados.get('linhaFator', '4'))


def get_coluna_fator(dados):
    return dados.get('colunaFator', "G")


def get_coluna_index(coluna):
    return column_index_from_string(coluna)


def get_valor_bdi_formatado(dados):
    return "{:.2%}".format(float(get_valor_bdi(dados))/100).replace('.', ',')

# composicao


def get_planilha_comp(dados):
    return dados.get('planilhaComposicao', 'Composicao')


def get_descricao_comp(dados):
    return dados.get('composicaoDescricao', 'A')


def get_item_descricao_comp_aux(dados):
    return dados.get('colunaItemDescricaoComposicao', 'B')


def get_coeficiente_comp(dados):
    return dados.get('composicaoCoeficiente', 'E')


def get_copiar_coeficiente_comp(dados):
    return dados.get(
        'composicaoCoeficienteCopiar', 'L')


def get_preco_unitario_comp(dados):
    return dados.get(
        'composicaoPrecoUnitario', 'F')


def get_copiar_preco_unitario_comp(dados):
    return dados.get(
        'composicaoPrecoUnitarioCopiar', 'M')


def get_coluna_totais_comp(dados):
    return dados.get(
        'colunaTotaisComposicao',
        'E'
    )


def get_valor_totais_comp(dados):
    return dados.get(
        'valorTotaisComposicao',
        'G'
    )

# AUXILIAR


def get_planilha_aux(dados):
    return dados.get('planilhaAuxiliar', 'Composicao Auxiliares')


def get_descricao_aux(dados):
    return dados.get('auxiliarDescricao', 'A')


def get_coeficiente_aux(dados):
    return dados.get(
        'auxiliarCoeficiente', 'E')


def get_copiar_coeficiente_aux(dados):
    return dados.get(
        'auxiliarCoeficienteCopiar', 'L')


def get_preco_unitario_aux(dados):
    return dados.get(
        'auxiliarPrecoUnitario', 'F')


def get_copiar_preco_unitario_aux(dados):
    return dados.get(
        'auxiliarPrecoUnitarioCopiar', 'M')


def get_coluna_totais_aux(dados):
    return dados.get(
        'colunaTotaisAuxiliar',
        'E'
    )


def get_valor_totais_aux(dados):
    return dados.get(
        'valorTotaisAuxiliar',
        'G'
    )


def get_valor_string(dados):
    return dados.get('valor', 'VALOR:')


def get_valor_com_bdi_string(dados):
    return dados.get('valorComBdi', 'VALOR COM BDI')
