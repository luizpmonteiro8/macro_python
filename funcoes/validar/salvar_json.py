"""
Componente para salvar dados no arquivo JSON de configuração.
"""

import json

CAMINHO_JSON = "config/valores_colunas.json"


def salvar_json_corrigido(dados, indice_config=0):
    """
    Salva os dados corrigidos no arquivo JSON.

    Args:
        dados: Lista de configurações atualizadas
        indice_config: Índice da configuração no array JSON
    """
    with open(CAMINHO_JSON, "w", encoding="utf-8") as f:
        json.dump(dados, f, indent=2, ensure_ascii=False)
