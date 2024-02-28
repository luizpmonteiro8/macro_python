# orcamentaria
def get_label_planilha_orcamentaria(valores):
    return valores['planilhaOrcamentaria'].get('planilhaOrcamentariaNome',
                                               'Planilha orcamentaria')


def get_title_planilha_orcamentaria(valores):
    return valores['planilhaOrcamentaria'].get('tituloPlanilhaOrcamentaria',
                                               'PLANILHA ORÇAMENTARIA')


def get_label_coluna_inicial(valores):
    return valores['planilhaOrcamentaria'].get('colunaicial',
                                               'Coluna inicial')


def get_label_coluna_final(valores):
    return valores['planilhaOrcamentaria'].get('colunaFinal',
                                               'Coluna final')


def get_label_valor_inicial(valores):
    return valores['planilhaOrcamentaria'].get('valorInicial',
                                               'Valor inicial')


def get_label_valor_final(valores):
    return valores['planilhaOrcamentaria'].get('valorFinal',
                                               'Valor final')


def get_label_preco_unitario_copiar(valores):
    return valores['planilhaOrcamentaria'].get('precoUnitarioCopiar',
                                               'Preço unitário copiar')

# fator


def get_title_planilha_fator(valores):
    return valores['planilhaFator'].get('tituloPlanilhaFator',
                                        'PLANILHA FATOR')


def get_label_planilha_fator(valores):
    return valores['planilhaFator'].get('planilhaFator', 'Planilha fator')


def get_label_BDI(valores):
    return valores['planilhaFator'].get('BDI', 'BDI')


def get_label_coluna_fator(valores):
    return valores['planilhaFator'].get('colunaFator', 'Coluna fator')


def get_label_linha_fator(valores):
    return valores['planilhaFator'].get('linhaFator', 'Linha fator')

# aux


def get_title_planilha_aux(valores):
    return valores['planilhaAuxiliar'].get('tituloPlanilhaAuxiliar',
                                           'PLANILHA AUXILIAR')


def get_label_planilha_aux(valores):
    return valores['planilhaAuxiliar'].get('planilhaAuxiliar',
                                           'Planilha auxiliar')


# composicao
def get_title_planilha_composicao(valores):
    return valores['planilhaComposicao'].get('tituloPlanilhaComposicao',
                                             'PLANILHA COMPOSICAO')


def get_label_planilha_composicao(valores):
    return valores['planilhaComposicao'].get('planilhaComposicao',
                                             'Composicoes')


def get_label_coluna_item(valores):
    return valores['planilhaComposicao'].get('colunaItem', 'Coluna item')


# composicao aux
def get_label_composicao_auxiliar_coeficiente_copiar(valores):
    return valores['planilhaComposicaoAuxiliar'].get(
        'coeficienteCopiar',
        'Coluna copiar coeficiente'
    )


def get_label_composicao_auxiliar_preco_unitario_copiar(valores):
    return valores['planilhaComposicaoAuxiliar'].get(
        'precoUnitarioCopiar',
        'Coluna copiar preço unitário'
    )


def get_label_composicao_auxiliar_coluna_totais(valores):
    return valores['planilhaComposicaoAuxiliar'].get(
        'colunaTotais',
        'Coluna totais'
    )


def get_label_composicao_auxiliar_valor_totais(valores):
    return valores['planilhaComposicaoAuxiliar'].get(
        'valorTotais',
        'Valor totais'
    )


# todos

def get_label_codigo(valores):
    return valores['todos'].get('codigo', 'Código')


def get_label_descricao(valores):
    return valores['todos'].get('descricao', 'Descricão')


def get_label_quantidade(valores):
    return valores['todos'].get('quantidade', 'Quantidade')


def get_label_coeficiente(valores):
    return valores['todos'].get('coeficiente', 'Coeficiente')


def get_label_preco_unitario(valores):
    return valores['todos'].get('precoUnitario', 'Preço unitário')


def get_label_preco_total(valores):
    return valores['todos'].get('precoTotal', 'Preço total')


def get_label_valor(valores):
    return valores['todos'].get('valor', 'Valor')


def get_label_valor_bdi(valores):
    return valores['todos'].get('valorComBdi', 'Valor com BDI')
