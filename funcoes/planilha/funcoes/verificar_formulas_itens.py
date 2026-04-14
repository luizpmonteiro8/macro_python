from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell


def verificar_e_adicionar_formulas(workbook, dados):
    """
    Verifica e adiciona fórmulas para itens com 'buscarAuxiliar': 'Sim'.
    Para células mescladas com código, busca a próxima linha com "VALOR:" para referência.
    """
    print(">>> Verificando fórmulas dos itens auxiliares...")

    sheet = workbook[dados.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")]
    col_item = dados.get("auxiliarDescricao", "A")  # Coluna do código
    col_preco = dados.get("auxiliarPrecoUnitario", "F")  # Coluna F
    val_str = dados.get("valor", "VALOR:")  # String "VALOR:"

    col_item_idx = column_index_from_string(col_item)
    col_preco_idx = column_index_from_string(col_preco)

    max_row = sheet.max_row
    linhas_modificadas = []

    # Construir lista de labels para pular dinamicamente do JSON
    # Coleta todos os nomes dos itens com buscarAuxiliar: "Sim"
    labels_pular = []
    totals_pular = []  # Lista de strings "TOTAL xxx:"
    if dados and isinstance(dados, list) and len(dados) > 0:
        primeiro_item = dados[0]
        for key, item in primeiro_item.items():
            if key.startswith("item") and isinstance(item, dict):
                if item.get("buscarAuxiliar") == "Sim":
                    nome = item.get("nome", "")
                    if nome:
                        labels_pular.append(nome.upper())
                    total_str = item.get("total", "")
                    if total_str:
                        totals_pular.append(total_str.upper())

    print(f">> Labels para pular (buscarAuxiliar: Sim): {labels_pular}")
    print(f">> Totals para pular: {totals_pular}")

    # Primeiro: construir mapa de códigos -> linha de célula mesclada + próxima linha com "VALOR:"
    mapa_codigos_valor = {}

    merged_ranges = list(sheet.merged_cells.ranges)
    for merged_range in merged_ranges:
        # Verificar se a célula mesclada está na coluna A
        if merged_range.min_col <= col_item_idx <= merged_range.max_col:
            cell_val = sheet.cell(row=merged_range.min_row, column=col_item_idx).value
            if cell_val:
                codigo = str(cell_val).strip()
                # Limpar caracteres especiais
                codigo_limpo = (
                    codigo.replace("\u200b", "").replace("\ufeff", "").strip()
                )
                # Extrair parte numérica (pode ter descrição depois)
                parte_numerica = "".join(c for c in codigo_limpo if c.isdigit())[:8]
                # Usar código completo limpo como chave também (para códigos como ED-5227)
                codigo_completo = codigo_limpo.replace(" ", "").upper()

                if parte_numerica or codigo_completo:
                    # Buscar próxima linha com "VALOR:" a partir da célula mesclada
                    # O "VALOR:" está na coluna E
                    linha_valor = -1
                    for j in range(merged_range.min_row + 1, max_row + 1):
                        cell_e = sheet.cell(row=j, column=5).value  # Coluna E
                        if cell_e and val_str.upper() in str(cell_e).upper():
                            linha_valor = j
                            break

                    if linha_valor > 0:
                        # Armazenar tanto por parte numérica quanto por código completo
                        if parte_numerica:
                            mapa_codigos_valor[parte_numerica] = linha_valor
                        # Também armazenar se começar com código como ED-, MOED-, etc
                        if (
                            len(codigo_completo) >= 5
                            and not codigo_completo[0].isdigit()
                        ):
                            mapa_codigos_valor[codigo_completo] = linha_valor

    print(f">> Códigos com referência a 'VALOR:': {len(mapa_codigos_valor)}")

    # Segundo: adicionar fórmulas para linhas que têm código mas sem valor em F
    for i in range(1, max_row + 1):
        cell_item = sheet.cell(row=i, column=col_item_idx).value
        if not cell_item or isinstance(cell_item, MergedCell):
            continue

        # Pular cabeçalhos usando labels dinâmicos do JSON
        codigo = str(cell_item).strip()
        codigo_upper = codigo.upper()
        if any(x in codigo_upper for x in labels_pular):
            continue
        if any(x in codigo_upper for x in totals_pular):
            continue

        cell_f = sheet.cell(row=i, column=col_preco_idx).value
        if isinstance(cell_f, MergedCell):
            continue

        # Se já tem fórmula, pular
        if cell_f and isinstance(cell_f, str) and cell_f.startswith("="):
            continue

        codigo_limpo = codigo.replace("\u200b", "").replace("\ufeff", "").strip()
        parte_numerica = "".join(c for c in codigo_limpo if c.isdigit())[:8]
        codigo_completo = codigo_limpo.replace(" ", "").upper()

        # Verificar tanto parte numérica quanto código completo (ex: ED-5227)
        chave_encontrada = None
        if (
            parte_numerica
            and len(parte_numerica) >= 5
            and parte_numerica in mapa_codigos_valor
        ):
            chave_encontrada = parte_numerica
        elif (
            len(codigo_completo) >= 5
            and not codigo_completo[0].isdigit()
            and codigo_completo in mapa_codigos_valor
        ):
            chave_encontrada = codigo_completo

        if chave_encontrada:
            ref_linha = mapa_codigos_valor[chave_encontrada]
            if ref_linha != i:
                cell_destino = sheet.cell(row=i, column=col_preco_idx)
                if not isinstance(cell_destino, MergedCell):
                    formula = f"=G{ref_linha}"
                    cell_destino.value = formula
                    linhas_modificadas.append((i, codigo_limpo[:50], ref_linha))

    print(f">> Fórmulas adicionadas: {len(linhas_modificadas)}")
    for linha, codigo, ref in linhas_modificadas[:10]:
        print(f"   L{linha}: {codigo} -> =G{ref}")
    if len(linhas_modificadas) > 10:
        print(f"   ... e mais {len(linhas_modificadas) - 10} fórmulas")

    return len(linhas_modificadas)
