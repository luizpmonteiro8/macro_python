from openpyxl.cell.cell import MergedCell

from .constantes import VALOR_LABEL
from .limpar import _limpar_codigo


def construir_mapa_mescladas(sheet, col_item_idx):
    """Constrói mapa de códigos para hyperlinks.

    Processa:
    1. Códigos em células mescladas grandes (>3 células)
    2. Códigos em células mescladas pequenas (<=3 células)
    3. Códigos em células normais (não mescladas)

    Retorna dict com {codigo: linha_valor}.
    """
    mapa_titulos = {}
    max_row = min(sheet.max_row, 20000)

    # Conjunto para rastrear linhas já processadas (evitar duplicatas)
    linhas_processadas = set()

    # Processar linha por linha
    for linha in range(1, max_row + 1):
        # Pular se já processou esta linha
        if linha in linhas_processadas:
            continue

        cell = sheet.cell(row=linha, column=col_item_idx)
        val = cell.value

        if not val:
            continue

        # Limpar o valor
        val_str = str(val).replace("\u200b", "").replace("\ufeff", "").strip()

        # Extrair código
        codigo = _limpar_codigo(val_str)
        if not codigo or len(codigo) < 5:
            continue

        codigo_upper = codigo.upper()

        # Verificar se está em célula mesclada
        merged_range = None
        for mr in sheet.merged_cells.ranges:
            if (
                mr.min_row <= linha <= mr.max_row
                and mr.min_col <= col_item_idx <= mr.max_col
            ):
                merged_range = mr
                break

        # Buscar linha do VALOR: abaixo do código
        linha_valor = buscar_linha_valor(sheet, linha, max_row)

        if merged_range:
            # Se está em célula mesclada, marcar todas as linhas como processadas
            for l in range(merged_range.min_row, merged_range.max_row + 1):
                linhas_processadas.add(l)

            # Incluir MESMO se for pequena (<=3 células)
            # (códigos em células normais também entram aqui)
            mapa_titulos[codigo_upper] = linha_valor
        else:
            # Célula não mesclada - adicionar ao mapa
            mapa_titulos[codigo_upper] = linha_valor
            linhas_processadas.add(linha)

    return mapa_titulos


def buscar_linha_valor(sheet, linha_inicio, max_row):
    """Busca linha com 'VALOR:' a partir da linha_inicio."""
    for linha in range(linha_inicio + 1, max_row + 1):
        val_e = sheet.cell(row=linha, column=5).value
        if val_e and VALOR_LABEL.upper() in str(val_e).upper():
            return linha
    return None
