"""
Sistema de validação de arquivos Excel e configurações JSON.
Componentes modulares para validação de planilhas.

Uso:
    from funcoes.validar import validar_arquivo_excel

    resultado = validar_arquivo_excel("caminho/para/arquivo.xlsx", dados)
"""

from funcoes.validar.janela_corrigir import janela_corrigir_valor
from funcoes.validar.salvar_json import salvar_json_corrigido
from funcoes.validar.validar_estrutura_base import validar_estrutura_base
from funcoes.validar.validar_nome_planilha import validar_nome_planilha
from funcoes.validar.validar_coluna import validar_coluna_existe
from funcoes.validar.validar_valor_celula import validar_valor_celula
from funcoes.validar.validar_celula_bdi import validar_celula_bdi
from funcoes.validar.validar_valor_existe_na_coluna import validar_valor_existe_na_coluna
from funcoes.validar.validar_planilha_orcamentaria import validar_planilha_orcamentaria
from funcoes.validar.validar_planilha_resumo import validar_planilha_resumo
from funcoes.validar.validar_planilha_composicoes import validar_planilha_composicoes
from funcoes.validar.validar_planilha_composicoes_auxiliares import (
    validar_planilha_composicoes_auxiliares,
)
from funcoes.validar.funcao_principal import (
    validar_arquivo_excel,
    validar_todas_configuracoes,
)

__all__ = [
    "janela_corrigir_valor",
    "salvar_json_corrigido",
    "validar_estrutura_base",
    "validar_nome_planilha",
    "validar_coluna_existe",
    "validar_valor_celula",
    "validar_celula_bdi",
    "validar_valor_existe_na_coluna",
    "validar_planilha_orcamentaria",
    "validar_planilha_resumo",
    "validar_planilha_composicoes",
    "validar_planilha_composicoes_auxiliares",
    "validar_arquivo_excel",
    "validar_todas_configuracoes",
]
