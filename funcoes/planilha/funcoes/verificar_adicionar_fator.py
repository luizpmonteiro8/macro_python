from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell


def verificar_e_adicionar_fator(workbook, dados):
    """
    Verifica e adiciona fórmulas de fator para itens com 'adicionarFator': 'Sim'.

    Lógica:
    - Se fatorCoeficiente: "Sim" → usar coluna E (composicaoCoeficiente ou auxiliarCoeficiente)
    - Se fatorCoeficiente: "Não" → usar coluna F (composicaoPrecoUnitario ou auxiliarPrecoUnitario)

    Valida em ambas as planilhas: composições e composições auxiliares.
    """
    print(">>> Verificando e adicionando fatores dos itens...")

    # O dados pode ser uma lista [{}] ou um dicionário {}
    dados_itens = dados
    if isinstance(dados, list) and len(dados) > 0:
        dados_itens = dados[0]

    # Obter nomes das planilhas do JSON
    planilha_comp = dados_itens.get("planilhaComposicao", "COMPOSICOES")
    planilha_aux = dados_itens.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")

    # Obter colunas do JSON
    col_coef_comp = dados_itens.get("composicaoCoeficiente", "E")
    col_preco_comp = dados_itens.get("composicaoPrecoUnitario", "F")
    col_coef_comp_antigo = dados_itens.get("composicaoCoeficienteCopiar", "L")
    col_preco_comp_antigo = dados_itens.get("composicaoPrecoUnitarioCopiar", "M")
    col_desc_comp = dados_itens.get("composicaoDescricao", "A")

    col_coef_aux = dados_itens.get("auxiliarCoeficiente", "E")
    col_preco_aux = dados_itens.get("auxiliarPrecoUnitario", "F")
    col_coef_aux_antigo = dados_itens.get("auxiliarCoeficienteCopiar", "L")
    col_preco_aux_antigo = dados_itens.get("auxiliarPrecoUnitarioCopiar", "M")
    col_desc_aux = dados_itens.get("auxiliarDescricao", "A")

    # Construir lista de itens que precisam de fator (MANTER INDIVIDUAL)
    # Cada item é processado separadamente com seus próprios filtros
    itens_fator = []

    for key, item in dados_itens.items():
        if key.startswith("item") and isinstance(item, dict):
            if item.get("adicionarFator") == "Sim":
                itens_fator.append(
                    {
                        "nome": item.get("nome", ""),
                        "total": item.get("total", ""),
                        "fatorCoeficiente": item.get("fatorCoeficiente") == "Sim",
                        "iniciaPor": item.get("iniciaPor", ""),
                        "naoIniciaPor": item.get("naoIniciaPor", ""),
                    }
                )

    print(f">> Itens com adicionarFator: Sim - {len(itens_fator)}")
    for item in itens_fator:
        print(
            f"   {item['nome']}: fatorCoeficiente={item['fatorCoeficiente']}, "
            f"iniciaPor='{item['iniciaPor']}', naoIniciaPor='{item['naoIniciaPor']}'"
        )

    # Labels para pular (itens com buscarAuxiliar: "Não")
    labels_pular = []
    totals_pular = []
    for key, item in dados_itens.items():
        if key.startswith("item") and isinstance(item, dict):
            if item.get("buscarAuxiliar") == "Não":
                nome = item.get("nome", "")
                if nome:
                    labels_pular.append(nome.upper())
                total_str = item.get("total", "")
                if total_str:
                    totals_pular.append(total_str.upper())

    print(f">> Labels para pular (buscarAuxiliar: Não): {labels_pular}")
    print(f">> Totals para pular: {totals_pular}")

    total_adicionados_comp = 0
    total_adicionados_aux = 0

    # Verificar composições
    print(f"\n>> Verificando planilha: {planilha_comp}")
    # ✅ NA PLANILHA PRINCIPAL COMPOSIÇÕES: PROCESSAR TODOS OS ITENS, INCLUSIVE OS COM buscarAuxiliar: Não
    adicionados = verificar_e_adicionar_planilha(
        workbook[planilha_comp],
        column_index_from_string(col_desc_comp),
        column_index_from_string(col_coef_comp),
        column_index_from_string(col_preco_comp),
        column_index_from_string(col_coef_comp_antigo),
        column_index_from_string(col_preco_comp_antigo),
        itens_fator,
        [],  # NAO PULAR NADA NA PLANILHA PRINCIPAL
        totals_pular,
    )
    total_adicionados_comp = adicionados
    print(f">> Fórmulas adicionadas em Composições: {adicionados}")

    # Verificar composições auxiliares
    print(f"\n>> Verificando planilha: {planilha_aux}")
    # ✅ PROCESSAR TODOS OS ITENS com adicionarFator: "Sim" NA PLANILHA AUXILIAR
    # O filtro buscarAuxiliar: "Não" é irrelevante para a aplicação de fórmulas de fator
    itens_fator_aux = itens_fator
    adicionados = verificar_e_adicionar_planilha(
        workbook[planilha_aux],
        column_index_from_string(col_desc_aux),
        column_index_from_string(col_coef_aux),
        column_index_from_string(col_preco_aux),
        column_index_from_string(col_coef_aux_antigo),
        column_index_from_string(col_preco_aux_antigo),
        itens_fator_aux,
        [],
        totals_pular,
    )
    total_adicionados_aux = adicionados
    print(f">> Fórmulas adicionadas em Composições Auxiliares: {adicionados}")

    # Resumo
    total = total_adicionados_comp + total_adicionados_aux
    print(f"\n>>> Total de fórmulas de fator adicionadas: {total}")

    return total


def encontrar_todas_secoes(sheet, col_desc_idx, nome_upper, total_upper, inicia_por="", nao_inicia_por=""):
    """
    Encontra TODAS as seções com o nome dado na planilha.
    Retorna lista de tuplas (inicio, fim) para cada seção.
    ONLY retorna seções onde o título começa com inicia_por (se fornecido)
    e NÃO começa com nao_inicia_por (se fornecido).
    """
    secoes = []
    colunas_busca = list(range(col_desc_idx, col_desc_idx + 10))
    max_row = sheet.max_row

    # Primeiro, encontrar todos os inícios que passam nos filtros
    inicios = []
    for i in range(1, max_row + 1):
        for col_busca in colunas_busca:
            cell = sheet.cell(row=i, column=col_busca).value
            if cell and isinstance(cell, MergedCell):
                continue
            cell_str = str(cell).strip()
            if cell and nome_upper in cell_str.upper():
                # Verificar que não é um TOTAL
                if "TOTAL" not in cell_str.upper():
                    # Se inicia_por fornecido, verificar que o título começa com ele
                    if inicia_por and not cell_str.startswith(inicia_por):
                        continue
                    # Se nao_inicia_por fornecido, verificar que o título NÃO começa com ele
                    if nao_inicia_por and cell_str.startswith(nao_inicia_por):
                        continue
                    inicios.append((i, col_busca))
                    break

    # Para cada início, encontrar o fim mais próximo
    for inicio, col_inicio in inicios:
        fim = -1
        for j in range(inicio + 1, max_row + 1):
            for col_busca in colunas_busca:
                cell = sheet.cell(row=j, column=col_busca).value
                if cell and total_upper in str(cell).upper():
                    fim = j
                    break
            if fim != -1:
                break
        if fim != -1:
            secoes.append((inicio, fim))

    return secoes


def verificar_e_adicionar_planilha(
    sheet,
    col_desc_idx,
    col_coef_idx,
    col_preco_idx,
    col_coef_antigo_idx,
    col_preco_antigo_idx,
    itens_fator,
    labels_pular,
    totals_pular,
):
    """
    Verifica uma planilha específica e adiciona fórmulas de fator onde necessário.
    Processa cada item em TODAS as seções correspondentes.
    """
    linhas_adicionadas = []
    max_row = sheet.max_row

    # Encontrar TODAS as seções para cada nome único
    # NÃO usar mais nomes_processados para permitir múltiplos itens com mesmo nome
    secoes_encontradas = {}  # cache_key (nome_upper, iniciaPor, naoIniciaPor) -> [(inicio, fim), ...]

    # Primeiro pass: encontrar todas as seções únicas
    for item_info in itens_fator:
        nome = item_info["nome"]
        total_str = item_info["total"]
        nome_upper = nome.upper()
        total_upper = total_str.upper() if total_str else ""
        inicia_por = item_info["iniciaPor"]
        nao_inicia_por = item_info["naoIniciaPor"]

        # Criar chave única que inclui os filtros iniciaPor e naoIniciaPor
        cache_key = (nome_upper, inicia_por, nao_inicia_por)

        # Se já buscamos essa seção com os mesmos filtros, não buscar de novo
        if cache_key in secoes_encontradas:
            continue

        # Buscar todas as seções com esse nome (passando inicia_por e nao_inicia_por para filtrar na busca)
        secoes = encontrar_todas_secoes(sheet, col_desc_idx, nome_upper, total_upper, inicia_por, nao_inicia_por)
        if secoes:
            secoes_encontradas[cache_key] = secoes
            print(f"   Encontradas {len(secoes)} seções de '{nome_upper}' (iniciaPor='{inicia_por}', naoIniciaPor='{nao_inicia_por}')")

    # Processar cada item em TODAS as seções correspondentes
    for item_info in itens_fator:
        nome = item_info["nome"]
        total_str = item_info["total"]
        fator_coef = item_info["fatorCoeficiente"]
        inicia_por = item_info["iniciaPor"]
        nao_inicia_por = item_info["naoIniciaPor"]

        nome_upper = nome.upper()
        cache_key = (nome_upper, inicia_por, nao_inicia_por)

        # Obter lista de seções para este nome
        if cache_key not in secoes_encontradas:
            print(f"   !! Não encontrou nenhuma seção para: {nome_upper} (iniciaPor='{inicia_por}', naoIniciaPor='{nao_inicia_por}')")
            continue

        secoes = secoes_encontradas[cache_key]
        count_processadas = 0

        # Processar TODAS as seções que correspondem ao filtro
        for idx, (inicio, fim) in enumerate(secoes):
            count_processadas += 1
            print(
                f"   Processando [{count_processadas}]: {nome_upper} (L{inicio} a L{fim}), "
                f"iniciaPor='{inicia_por}', naoIniciaPor='{nao_inicia_por}'"
            )

            # Processar linhas entre início e fim
            for y in range(inicio + 1, fim):
                cell_desc = sheet.cell(row=y, column=col_desc_idx).value
                if isinstance(cell_desc, MergedCell):
                    continue

                desc = str(cell_desc) if cell_desc else ""

                # Pular linhas de títulos de seção (que contêm palavras como FONTE, UNID, COEFICIENTE)
                desc_upper = desc.upper()
                if "COEFICIENTE" in desc_upper or "PREÇO UNITÁRIO" in desc_upper:
                    continue

                # Verificar se é uma linha de título de seção (contém FONTE ou UNID ou TOTAL)
                cell_fonte = sheet.cell(row=y, column=col_desc_idx + 2).value
                cell_unid = sheet.cell(row=y, column=col_desc_idx + 3).value
                if cell_fonte and "FONTE" in str(cell_fonte).upper():
                    continue
                if cell_unid and "UNID" in str(cell_unid).upper():
                    continue

                # NÃO pular linhas só porque a seção tem um nome em labels_pular
                # O labels_pular serve para outros contextos, não para pular a seção inteira
                # Apenas pular linhas que são especificamente o label (título da seção)
                # e linhas de TOTAL
                # ✅ CORREÇÃO: labels_pular NÃO deve pular linhas INTERNAS da seção
                # Removida verificação que pulava TODAS as linhas de itens com buscarAuxiliar: "Não"

                # Apenas pular linhas de TOTAL
                if any(x in desc_upper for x in totals_pular):
                    continue

                # Adicionar fórmula
                if fator_coef:
                    # Adicionar na coluna E (coeficiente)
                    cell_coef = sheet.cell(row=y, column=col_coef_idx)
                    if not isinstance(cell_coef, MergedCell):
                        # SOBRESCREVER SEMPRE quando fatorCoeficiente=True
                        # Isso garante que valores numéricos sejam substituídos
                        formula = f"={get_column_letter(col_coef_antigo_idx)}{y}*FATOR"
                        cell_coef.value = formula
                        linhas_adicionadas.append(f"L{y}: {desc[:30]} -> {formula}")
                else:
                    # Adicionar na coluna F (preço unitário)
                    cell_preco = sheet.cell(row=y, column=col_preco_idx)
                    if not isinstance(cell_preco, MergedCell):
                        if cell_preco.value is None or (
                            not isinstance(cell_preco.value, str)
                            or not cell_preco.value.startswith("=")
                        ):
                            formula = f"=ROUND({get_column_letter(col_preco_antigo_idx)}{y}*FATOR, 2)"
                            cell_preco.value = formula
                            linhas_adicionadas.append(f"L{y}: {desc[:30]} -> {formula}")

        if count_processadas > 0:
            print(f"   Processadas {count_processadas} seções de '{nome_upper}'")

    # Mostrar primeiras linhas adicionadas
    if linhas_adicionadas:
        print(f"   Adicionadas: {len(linhas_adicionadas)}")
        for item in linhas_adicionadas[:10]:
            print(f"      {item}")
        if len(linhas_adicionadas) > 10:
            print(f"      ... e mais {len(linhas_adicionadas) - 10}")

    return len(linhas_adicionadas)
