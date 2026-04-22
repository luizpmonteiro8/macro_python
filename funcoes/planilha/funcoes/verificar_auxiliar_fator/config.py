"""Extrai configurações do valores_item.json."""


def extrair_configuracoes(todos_item):
    """Extrai mapa_nome_inicia e mapa_config do valores_item.json.

    Args:
        todos_item: Lista ou dict com itens do JSON

    Returns:
        tuple: (mapa_nome_inicia, mapa_config)
    """
    todos_item_dict = todos_item[0] if isinstance(todos_item, list) else todos_item
    todos_item_data = []
    for key, value in todos_item_dict.items():
        if key.startswith("item") and isinstance(value, dict):
            todos_item_data.append(value)

    mapa_nome_inicia = []
    mapa_config = []

    for item in todos_item_data:
        if not isinstance(item, dict):
            continue
        nome = item.get("nome", "")
        if not nome:
            continue

        mapa_nome_inicia.append(
            {
                "nome": nome,
                "iniciaPor": item.get("iniciaPor", ""),
                "naoIniciaPor": item.get("naoIniciaPor", ""),
            }
        )

        mapa_config.append(
            {
                "nome": nome,
                "total": item.get("total", ""),
                "adicionarFator": item.get("adicionarFator", "Não"),
                "buscarAuxiliar": item.get("buscarAuxiliar", "Não"),
                "fatorCoeficiente": item.get("fatorCoeficiente", "Não") == "Sim",
            }
        )

    return mapa_nome_inicia, mapa_config
