from openpyxl.cell.cell import MergedCell

from .constantes import VALOR_LABEL
from .limpar import _limpar_codigo


def _construir_mapa_mescladas(sheet, col_item_idx):
    """Constrói mapa de códigos para hyperlinks usando APENAS células mescladas
    com mais de 3 células na coluna desejada. Retorna dict com
    {codigo: linha_valor}.

    Percorre linha por linha e avança para encontrar VALOR: correspondente.
    """
    mapa_titulos = {}
    max_row = min(sheet.max_row, 20000)

    # Processar linha por linha
    for linha in range(1, max_row + 1):
        cell = sheet.cell(row=linha, column=col_item_idx)

        # Verificar se a célula é parte de uma célula mesclada
        if isinstance(cell, MergedCell):
            # Encontrar o intervalo mesclado ao qual esta célula pertence
            merged_range = None
            for mr in sheet.merged_cells.ranges:
                if (
                    mr.min_row <= linha <= mr.max_row
                    and mr.min_col <= col_item_idx <= mr.max_col
                ):
                    merged_range = mr
                    break

            # Se não encontrou intervalo mesclado, continuar
            if not merged_range:
                continue

            # Verificar se esta célula é a primeira do intervalo (min_row)
            if linha != merged_range.min_row:
                continue

            # Filtrar apenas mesclagens maiores que 3 células
            total_celulas = (merged_range.max_row - merged_range.min_row + 1) * (
                merged_range.max_col - merged_range.min_col + 1
            )
            if total_celulas <= 3:
                continue

            # Obter o valor do título
            val = sheet.cell(row=merged_range.min_row, column=col_item_idx).value

            if val:
                codigo = _limpar_codigo(str(val))
                if codigo and len(codigo) >= 5:
                    linha_titulo = merged_range.min_row
                    linha_valor = None

                    # Avançar o for até encontrar VALOR:
                    for linha_busca in range(linha_titulo + 1, max_row + 1):
                        val_e = sheet.cell(row=linha_busca, column=5).value
                        if val_e and VALOR_LABEL.upper() in str(val_e).upper():
                            linha_valor = linha_busca
                            break

                    mapa_titulos[codigo.upper()] = linha_valor

    return mapa_titulos
