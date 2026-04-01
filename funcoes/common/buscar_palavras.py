import re
import unicodedata


class CampoNaoEncontradoError(Exception):
    """Exceção levantada quando um campo não é encontrado na planilha."""

    def __init__(self, campo, valor_buscado="", contexto=""):
        self.campo = campo
        self.valor_buscado = valor_buscado
        self.contexto = contexto
        mensagem = f"❌ Campo não encontrado: '{campo}'"
        if valor_buscado:
            mensagem += f"\n   Valor que está sendo buscado: '{valor_buscado}'"
            mensagem += f"\n   ➡️ Vá na aba 'COMPOSICOES' e procure este texto na coluna de totais."
            mensagem += (
                f"\n   ➡️ Depois vá em 'Salvar dados' e altere o campo correspondente."
            )
        if contexto:
            mensagem += f"\n   {contexto}"
        super().__init__(mensagem)


def normalizar_texto(texto):
    if texto is None:
        return ""

    # normaliza unicode (remove NBSP, acentos estranhos, etc)
    texto = unicodedata.normalize("NFKD", str(texto))

    # troca NBSP por espaço normal
    texto = texto.replace("\xa0", " ")

    # remove conteúdo entre parênteses
    texto = re.sub(r"\(.*?\)", "", texto)

    # remove quebras, tabs e múltiplos espaços
    texto = re.sub(r"\s+", " ", texto)

    return texto.strip().lower()


def buscar_palavra(sheet, coluna, palavra):
    # Converte a letra da coluna para o número correspondente (A=1, B=2, etc.)
    numero_coluna = ord(coluna.upper()) - ord("A") + 1

    # Percorre as células da coluna e verifica se a palavra está presente
    for linha in range(1, sheet.max_row + 1):
        valor_celula = sheet.cell(row=linha, column=numero_coluna).value
        if valor_celula is not None:
            if palavra.lower() in str(valor_celula).lower():
                return linha

    # Se a palavra não foi encontrada, retorna algum valor indicativo, como -1
    return -1


def buscar_palavra_exata(sheet, coluna, palavra):
    # Converte a letra da coluna para o número correspondente (A=1, B=2, etc.)
    numero_coluna = ord(coluna.upper()) - ord("A") + 1

    # Percorre as células da coluna e verifica se a palavra está presente
    for linha in range(1, sheet.max_row + 1):
        valor_celula = sheet.cell(row=linha, column=numero_coluna).value
        if valor_celula is not None:
            if palavra.lower() == str(valor_celula).lower():
                return linha

    # Se a palavra não foi encontrada, retorna algum valor indicativo, como -1
    return -1


def buscar_palavra_com_linha(sheet, coluna, palavra, linha_inicial, linha_final):
    # Converte a letra da coluna para o número correspondente (A=1, B=2, etc.)
    numero_coluna = ord(coluna.upper()) - ord("A") + 1

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


def buscar_palavra_com_linha_exato(sheet, coluna, palavra, linha_inicial, linha_final):
    # Converte a letra da coluna para o número correspondente (A=1, B=2, etc.)
    numero_coluna = ord(coluna.upper()) - ord("A") + 1

    # Define a linha final como a última linha se não for fornecida
    linha_final = linha_final or sheet.max_row

    # Percorre as células da coluna e verifica se a palavra está presente
    for linha in range(linha_inicial, linha_final):
        valor_celula = sheet.cell(row=linha, column=numero_coluna).value
        if valor_celula is not None:
            if palavra.lower() == str(valor_celula).lower():
                return linha

    # Se a palavra não foi encontrada, retorna algum valor indicativo, como -1
    return -1


def buscar_palavra_com_linha_iniciando(
    sheet, coluna, palavra, linha_inicial, linha_final
):
    # Converte a letra da coluna para o número correspondente (A=1, B=2, etc.)
    numero_coluna = ord(coluna.upper()) - ord("A") + 1

    # Define a linha final como a última linha se não for fornecida
    linha_final = linha_final or sheet.max_row

    # Normaliza a palavra buscada (remove espaços extras e converte para minúsculo)
    palavra_normalizada = re.sub(r"\s+", " ", palavra.strip().lower())

    for linha in range(linha_inicial, linha_final + 1):
        valor_celula = sheet.cell(row=linha, column=numero_coluna).value
        if valor_celula is not None:
            texto_celula = re.sub(r"\s+", " ", str(valor_celula).strip().lower())
            if texto_celula.startswith(palavra_normalizada):
                return linha

    return -1


def buscar_palavra_contem(sheet, coluna, texto, lin_ini, lin_fim):
    numero_coluna = ord(coluna.upper()) - ord("A") + 1
    lin_fim = lin_fim or sheet.max_row
    print(">>> linha final:", lin_fim)
    texto_normalizado = normalizar_texto(texto)
    print(">>> Texto normalizado:", texto_normalizado)

    for linha in range(lin_ini, lin_fim + 1):
        valor_celula = sheet.cell(row=linha, column=numero_coluna).value

        if valor_celula:
            celula_normalizada = normalizar_texto(valor_celula)
            if texto_normalizado in celula_normalizada:
                return linha
    return -1
