import json


def open_valores_colunas():
    with open("config/valores_colunas.json", "r", encoding="utf-8") as f:
        return json.load(f)


def open_valores_label():
    with open("config/valores_label.json", "r", encoding="utf-8") as f:
        return json.load(f)


def open_valores_item():
    with open("config/valores_item.json", "r", encoding="utf-8") as f:
        return json.load(f)
