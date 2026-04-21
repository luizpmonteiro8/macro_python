from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell
from openpyxl.worksheet.hyperlink import Hyperlink


def verificar_auxiliar_fator(workbook, dados):
    """
    Função unificada que verifica e adiciona:
    1. Fórmulas para itens com 'buscarAuxiliar': 'Sim' (hiperlinks e referências)
    2. Fórmulas de fator para itens com 'adicionarFator': 'Sim'

    Lógica:
    - adicionarFator: "Sim":
        - fatorCoeficiente: "Sim" → adiciona fórmula no coeficiente (coluna E)
        - fatorCoeficiente: "Não" → adiciona fórmula no preço unitário (coluna F)
    - buscarAuxiliar: "Sim" → faz busca e adiciona hyperlink
    """
    print(">>> Iniciando verificação unificada de auxiliar e fator...")

    # O dados pode ser uma lista [{}] ou um dicionário {}
    dados_itens = dados
    if isinstance(dados, list) and len(dados) > 0:
        dados_itens = dados[0]

    # Obter nomes das planilhas do JSON
    planilha_comp = dados_itens.get("planilhaComposicao", "COMPOSICOES")
    planilha_aux = dados_itens.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")

    # Obter colunas do JSON
    col_desc_comp = dados_itens.get("composicaoDescricao", "A")
    col_coef_comp = dados_itens.get("composicaoCoeficiente", "E")
    col_preco_comp = dados_itens.get("composicaoPrecoUnitario", "F")
    col_coef_comp_antigo = dados_itens.get("composicaoCoeficienteCopiar", "L")
    col_preco_comp_antigo = dados_itens.get("composicaoPrecoUnitarioCopiar", "M")

    col_desc_aux = dados_itens.get("auxiliarDescricao", "A")
    col_coef_aux = dados_itens.get("auxiliarCoeficiente", "E")
    col_preco_aux = dados_itens.get("auxiliarPrecoUnitario", "F")
    col_coef_aux_antigo = dados_itens.get("auxiliarCoeficienteCopiar", "L")
    col_preco_aux_antigo = dados_itens.get("auxiliarPrecoUnitarioCopiar", "M")

    # Obter valores de linha a pular
    valor_label = dados_itens.get("valor", "VALOR:")
    valor_bdi_label = dados_itens.get("valorBdi", "VALOR BDI")
    valor_com_bdi_label = dados_itens.get("valorComBdi", "VALOR COM BDI")

    valores_pular = [
        valor_label.upper(),
        valor_bdi_label.upper(),
        valor_com_bdi_label.upper(),
    ]

    # ============================================
    # Construir lista de itens para buscarAuxiliar
    # ============================================
    itens_auxiliares = []
    for key, item in dados_itens.items():
        if key.startswith("item") and isinstance(item, dict):
            if item.get("buscarAuxiliar") == "Sim":
                itens_auxiliares.append(
                    {
                        "nome": item.get("nome", ""),
                        "total": item.get("total", ""),
                        "iniciaPor": item.get("iniciaPor", ""),
                        "naoIniciaPor": item.get("naoIniciaPor", ""),
                    }
                )

    print(f">> Itens com buscarAuxiliar: Sim - {len(itens_auxiliares)}")

    # ============================================
    # Construir lista de itens para adicionarFator
    # ============================================
    itens_fator = []
    for key, item in dados_itens.items():
        if key.startswith("item") and isinstance(item, dict):
            # Pular itens que têm buscarAuxiliar: "Sim" - eles recebem fórmula de referência
            if item.get("buscarAuxiliar") == "Sim":
                continue
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

    # ============================================
    # Construir listas de labels e totals para pular
    # ============================================
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

    total_resultados = {
        "formulas_auxiliares_comp": 0,
        "formulas_auxiliares_aux": 0,
        "formulas_fator_comp": 0,
        "formulas_fator_aux": 0,
        "hyperlinks_criados": 0,
    }

    # ============================================
    # Processar planilha COMP (Composições)
    # ============================================
    print(f"\n>> Processando planilha: {planilha_comp}")
    sheet_comp = workbook[planilha_comp]

    # 1. Verificar fórmulas auxiliares (buscarAuxiliar)
    formulas_aux = verificar_formulas_auxiliares_comp(
        sheet_comp,
        column_index_from_string(col_desc_comp),
        column_index_from_string(col_preco_comp),
        itens_auxiliares,
        planilha_aux,
    )
    total_resultados["formulas_auxiliares_comp"] = formulas_aux["formulas"]
    total_resultados["hyperlinks_criados"] += formulas_aux["hyperlinks"]

    # 2. Verificar e adicionar fator
    formulas_fator = verificar_fator_comp(
        sheet_comp,
        column_index_from_string(col_desc_comp),
        column_index_from_string(col_coef_comp),
        column_index_from_string(col_preco_comp),
        column_index_from_string(col_coef_comp_antigo),
        column_index_from_string(col_preco_comp_antigo),
        itens_fator,
        totals_pular,
        valores_pular,
    )
    total_resultados["formulas_fator_comp"] = formulas_fator

    print(
        f">> Composições: {formulas_aux['formulas']} fórmulas auxiliares, {formulas_fator} fórmulas de fator"
    )

    # ============================================
    # Processar planilha AUXILIAR
    # ============================================
    print(f"\n>> Processando planilha: {planilha_aux}")
    sheet_aux = workbook[planilha_aux]

    # 1. Verificar fórmulas auxiliares (buscarAuxiliar)
    formulas_aux = verificar_formulas_auxiliares_aux(
        sheet_aux,
        column_index_from_string(col_desc_aux),
        column_index_from_string(col_preco_aux),
        itens_auxiliares,
        planilha_aux,
        valor_label,
    )
    total_resultados["formulas_auxiliares_aux"] = formulas_aux["formulas"]
    total_resultados["hyperlinks_criados"] += formulas_aux["hyperlinks"]

    # 2. Verificar e adicionar fator
    formulas_fator = verificar_fator_aux(
        sheet_aux,
        column_index_from_string(col_desc_aux),
        column_index_from_string(col_coef_aux),
        column_index_from_string(col_preco_aux),
        column_index_from_string(col_coef_aux_antigo),
        column_index_from_string(col_preco_aux_antigo),
        itens_fator,
        totals_pular,
        valores_pular,
    )
    total_resultados["formulas_fator_aux"] = formulas_fator

    print(
        f">> Auxiliares: {formulas_aux['formulas']} fórmulas auxiliares, {formulas_fator} fórmulas de fator"
    )

    # ============================================
    # Resumo final
    # ============================================
    total_formulas = (
        total_resultados["formulas_auxiliares_comp"]
        + total_resultados["formulas_auxiliares_aux"]
        + total_resultados["formulas_fator_comp"]
        + total_resultados["formulas_fator_aux"]
    )

    print(f"\n>>> RESUMO:")
    print(f">> Total fórmulas adicionadas: {total_formulas}")
    print(f">> Total hyperlinks criados: {total_resultados['hyperlinks_criados']}")

    return total_resultados


# ============================================
# FUNÇÕES PARA FÓRMULAS AUXILIARES (buscarAuxiliar)
# ============================================


def verificar_formulas_auxiliares_comp(
    sheet, col_item_idx, col_preco_idx, itens_auxiliares, planilha_aux
):
    """
    Verifica e adiciona fórmulas para itens com buscarAuxiliar: "Sim" na planilha de composições.
    Cria hyperlinks na coluna de descrição (B) apontando para o título da seção.
    """
    print(">>> Verificando fórmulas dos itens auxiliares (Composições)...")

    col_desc_idx = 2  # Coluna B - descrição
    max_row = sheet.max_row

    linhas_modificadas = []
    hyperlinks_criados = 0

    # Construir mapa de códigos -> linha do título
    mapa_codigos_titulo = {}

    merged_ranges = list(sheet.merged_cells.ranges)
    for merged_range in merged_ranges:
        if merged_range.min_col <= col_item_idx <= merged_range.max_col:
            cell_val = sheet.cell(row=merged_range.min_row, column=col_item_idx).value
            if cell_val:
                codigo = str(cell_val).strip()
                codigo_limpo = (
                    codigo.replace("\u200b", "").replace("\ufeff", "").strip()
                )
                codigo_limpo = codigo_limpo.split()[0] if codigo_limpo.split() else ""
                codigo_completo = codigo_limpo.upper()

                if len(codigo_completo) >= 5:
                    mapa_codigos_titulo[codigo_completo] = merged_range.min_row

    # Processar cada linha
    for i in range(1, max_row + 1):
        cell_item = sheet.cell(row=i, column=col_item_idx).value
        if not cell_item or isinstance(cell_item, MergedCell):
            continue

        codigo = str(cell_item).strip()

        # Pular textos que não são códigos
        if any(
            x in codigo.upper()
            for x in [
                "MATERIAL",
                "MÃO DE OBRA",
                "SERVIÇO",
                "EQUIPAMENTO",
                "TOTAL",
                "PREÇO",
                "ENCARGOS",
            ]
        ):
            continue

        cell_f = sheet.cell(row=i, column=col_preco_idx).value
        if isinstance(cell_f, MergedCell):
            continue

        tem_formula = cell_f and isinstance(cell_f, str) and cell_f.startswith("=")

        codigo_limpo = codigo.replace("\u200b", "").replace("\ufeff", "").strip()
        codigo_limpo = codigo_limpo.split()[0] if codigo_limpo.split() else ""
        codigo_completo = codigo_limpo.upper()

        # Criar hyperlink para itens com fórmula existente
        if tem_formula:
            codigo_valido = codigo_completo and len(codigo_completo) >= 5
            linha_titulo = -1
            if codigo_valido and codigo_completo in mapa_codigos_titulo:
                linha_titulo = mapa_codigos_titulo[codigo_completo]

            if linha_titulo > 0:
                cell_desc = sheet.cell(row=i, column=col_desc_idx)
                if not isinstance(cell_desc, MergedCell) and not cell_desc.hyperlink:
                    cell_desc.hyperlink = Hyperlink(
                        ref=cell_desc.coordinate,
                        location=f"'{planilha_aux}'!A{linha_titulo}",
                    )
                    hyperlinks_criados += 1
            continue

        # Verificar se o código corresponde a algum filtro
        codigo_aprovado = verificar_filtros_codigo(codigo_completo, itens_auxiliares)
        if not codigo_aprovado:
            continue

        # Criar hyperlink
        if (
            codigo_completo
            and len(codigo_completo) >= 5
            and codigo_completo in mapa_codigos_titulo
        ):
            linha_titulo = mapa_codigos_titulo[codigo_completo]
            if linha_titulo > 0:
                cell_desc = sheet.cell(row=i, column=col_desc_idx)
                if not isinstance(cell_desc, MergedCell):
                    cell_desc.hyperlink = Hyperlink(
                        ref=cell_desc.coordinate,
                        location=f"'{planilha_aux}'!A{linha_titulo}",
                    )
                    hyperlinks_criados += 1

    print(f">> Fórmulas auxiliares em Composições: {len(linhas_modificadas)}")
    print(f">> Hyperlinks criados em Composições: {hyperlinks_criados}")

    return {"formulas": len(linhas_modificadas), "hyperlinks": hyperlinks_criados}


def verificar_formulas_auxiliares_aux(
    sheet, col_item_idx, col_preco_idx, itens_auxiliares, planilha_aux, val_str
):
    """
    Verifica e adiciona fórmulas para itens com buscarAuxiliar: "Sim" na planilha auxiliar.
    Aplica filtros iniciaPor/naoIniciaPor ao CÓDIGO do item.
    Também cria hyperlinks na coluna de descrição (B) apontando para o título da seção.
    """
    print(">>> Verificando fórmulas dos itens auxiliares (Auxiliares)...")

    col_desc_idx = 2  # Coluna B - descrição
    max_row = sheet.max_row

    linhas_modificadas = []
    hyperlinks_criados = 0

    # ============================================
    # Construir mapa de códigos -> linha do título (célula mesclada)
    # ============================================
    mapa_codigos_titulo = {}

    merged_ranges = list(sheet.merged_cells.ranges)
    for merged_range in merged_ranges:
        if merged_range.min_col <= col_item_idx <= merged_range.max_col:
            cell_val = sheet.cell(row=merged_range.min_row, column=col_item_idx).value
            if cell_val:
                codigo = str(cell_val).strip()
                codigo_limpo = (
                    codigo.replace("\u200b", "").replace("\ufeff", "").strip()
                )
                codigo_limpo = codigo_limpo.split()[0] if codigo_limpo.split() else ""
                codigo_completo = codigo_limpo.upper()

                if len(codigo_completo) >= 5:
                    mapa_codigos_titulo[codigo_completo] = merged_range.min_row

    # ============================================
    # Construir mapa de códigos -> linha de "VALOR:"
    # ============================================
    mapa_codigos_valor = {}

    for merged_range in merged_ranges:
        if merged_range.min_col <= col_item_idx <= merged_range.max_col:
            cell_val = sheet.cell(row=merged_range.min_row, column=col_item_idx).value
            if cell_val:
                codigo = str(cell_val).strip()
                codigo_limpo = (
                    codigo.replace("\u200b", "").replace("\ufeff", "").strip()
                )
                codigo_limpo = codigo_limpo.split()[0] if codigo_limpo.split() else ""
                codigo_completo = codigo_limpo.upper()

                if len(codigo_completo) >= 5:
                    linha_valor = -1
                    for j in range(merged_range.min_row + 1, max_row + 1):
                        cell_e = sheet.cell(row=j, column=5).value  # Coluna E
                        if cell_e and val_str.upper() in str(cell_e).upper():
                            linha_valor = j
                            break

                    if linha_valor > 0:
                        mapa_codigos_valor[codigo_completo] = linha_valor

    print(f">> Códigos com referência a 'VALOR:': {len(mapa_codigos_valor)}")

    # ============================================
    # Processar cada linha
    # ============================================
    for i in range(1, max_row + 1):
        cell_item = sheet.cell(row=i, column=col_item_idx).value
        if not cell_item or isinstance(cell_item, MergedCell):
            continue

        codigo = str(cell_item).strip()

        # Pular textos que não são códigos
        if any(
            x in codigo.upper()
            for x in [
                "MATERIAL",
                "MÃO DE OBRA",
                "SERVIÇO",
                "EQUIPAMENTO",
                "TOTAL",
                "PREÇO",
                "ENCARGOS",
            ]
        ):
            continue

        cell_f = sheet.cell(row=i, column=col_preco_idx).value
        if isinstance(cell_f, MergedCell):
            continue

        tem_formula = cell_f and isinstance(cell_f, str) and cell_f.startswith("=")

        codigo_limpo = codigo.replace("\u200b", "").replace("\ufeff", "").strip()
        codigo_limpo = codigo_limpo.split()[0] if codigo_limpo.split() else ""
        codigo_completo = codigo_limpo.upper()

        # Criar hyperlink para itens com fórmula existente
        if tem_formula:
            codigo_valido = codigo_completo and len(codigo_completo) >= 5
            linha_titulo = -1
            if codigo_valido and codigo_completo in mapa_codigos_titulo:
                linha_titulo = mapa_codigos_titulo[codigo_completo]

            if linha_titulo > 0:
                cell_desc = sheet.cell(row=i, column=col_desc_idx)
                if not isinstance(cell_desc, MergedCell) and not cell_desc.hyperlink:
                    cell_desc.hyperlink = Hyperlink(
                        ref=cell_desc.coordinate,
                        location=f"'{planilha_aux}'!A{linha_titulo}",
                    )
                    hyperlinks_criados += 1
            continue

        # Verificar se o código corresponde a algum filtro
        codigo_aprovado = verificar_filtros_codigo(codigo_completo, itens_auxiliares)
        if not codigo_aprovado:
            continue

        # Adicionar fórmula e hyperlink
        if (
            codigo_completo
            and len(codigo_completo) >= 5
            and codigo_completo in mapa_codigos_valor
        ):
            ref_linha = mapa_codigos_valor[codigo_completo]
            if ref_linha != i:
                cell_destino = sheet.cell(row=i, column=col_preco_idx)
                if not isinstance(cell_destino, MergedCell):
                    cell_destino.value = f"=G{ref_linha}"
                    linhas_modificadas.append((i, codigo_limpo[:50], ref_linha))

                    # Criar hyperlink
                    linha_titulo = -1
                    if codigo_completo in mapa_codigos_titulo:
                        linha_titulo = mapa_codigos_titulo[codigo_completo]

                    if linha_titulo > 0:
                        cell_desc = sheet.cell(row=i, column=col_desc_idx)
                        if not isinstance(cell_desc, MergedCell):
                            cell_desc.hyperlink = Hyperlink(
                                ref=cell_desc.coordinate,
                                location=f"'{planilha_aux}'!A{linha_titulo}",
                            )
                            hyperlinks_criados += 1

    print(f">> Fórmulas adicionadas em Auxiliares: {len(linhas_modificadas)}")
    print(f">> Hyperlinks criados em Auxiliares: {hyperlinks_criados}")

    return {"formulas": len(linhas_modificadas), "hyperlinks": hyperlinks_criados}


def verificar_filtros_codigo(codigo_upper, itens_auxiliares):
    """
    Verifica se o código corresponde a algum filtro dos itens auxiliares.
    Retorna True se o código for aprovado por algum item.
    """
    if not codigo_upper:
        return False

    # Primeiro: verificar se algum item tem filtros específicos definidos
    algum_filtro_especifico = any(
        item.get("iniciaPor") or item.get("naoIniciaPor") for item in itens_auxiliares
    )

    if algum_filtro_especifico:
        # Se algum item tem filtro específico, verificar se este código corresponde
        for item_info in itens_auxiliares:
            inicia_por = item_info.get("iniciaPor", "")
            nao_inicia_por = item_info.get("naoIniciaPor", "")

            # Ignorar itens com filtros vazios
            if not inicia_por and not nao_inicia_por:
                continue

            inicia_ok = True
            if inicia_por and not codigo_upper.startswith(inicia_por.upper()):
                inicia_ok = False

            nao_ok = True
            if nao_inicia_por and codigo_upper.startswith(nao_inicia_por.upper()):
                nao_ok = False

            if inicia_ok and nao_ok:
                return True

        return False
    else:
        # Se nenhum item tem filtro específico, aprovar todos
        return True


# ============================================
# FUNÇÕES PARA FATOR (adicionarFator)
# ============================================


def verificar_fator_comp(
    sheet,
    col_desc_idx,
    col_coef_idx,
    col_preco_idx,
    col_coef_antigo_idx,
    col_preco_antigo_idx,
    itens_fator,
    totals_pular,
    valores_pular,
):
    """
    Verifica e adiciona fórmulas de fator na planilha de composições.
    """
    print(">>> Verificando e adicionando fatores (Composições)...")

    linhas_adicionadas = []
    max_row = sheet.max_row

    # Encontrar todas as seções para cada nome único
    secoes_encontradas = {}
    nomes_processados = set()

    for item_info in itens_fator:
        nome = item_info["nome"]
        total_str = item_info["total"]
        nome_upper = nome.upper()
        total_upper = total_str.upper() if total_str else ""

        cache_key = nome_upper
        if cache_key in nomes_processados:
            continue

        nomes_processados.add(cache_key)

        secoes = encontrar_todas_secoes(
            sheet, col_desc_idx, nome_upper, total_upper, "", ""
        )
        if secoes:
            secoes_encontradas[cache_key] = secoes
            print(f"   Encontradas {len(secoes)} seções de '{nome_upper}'")

    # Processar cada item em todas as seções
    for item_info in itens_fator:
        nome = item_info["nome"]
        fator_coef = item_info["fatorCoeficiente"]
        inicia_por = item_info["iniciaPor"]
        nao_inicia_por = item_info["naoIniciaPor"]

        nome_upper = nome.upper()
        cache_key = nome_upper

        if cache_key not in secoes_encontradas:
            print(f"   !! Não encontrou nenhuma seção para: {nome_upper}")
            continue

        secoes = secoes_encontradas[cache_key]

        for idx, (inicio, fim) in enumerate(secoes):
            # Processar linhas entre início e fim
            for y in range(inicio + 1, fim):
                cell_desc = sheet.cell(row=y, column=col_desc_idx).value
                if isinstance(cell_desc, MergedCell):
                    continue

                desc = str(cell_desc) if cell_desc else ""
                desc_upper = desc.upper()

                # Pular linhas de títulos de seção
                if "COEFICIENTE" in desc_upper or "PREÇO UNITÁRIO" in desc_upper:
                    continue

                cell_fonte = sheet.cell(row=y, column=col_desc_idx + 2).value
                cell_unid = sheet.cell(row=y, column=col_desc_idx + 3).value
                if cell_fonte and "FONTE" in str(cell_fonte).upper():
                    continue
                if cell_unid and "UNID" in str(cell_unid).upper():
                    continue

                # Pular linhas de TOTAL
                if any(x in desc_upper for x in totals_pular):
                    continue
                if any(
                    x in desc_upper
                    for x in ["TOTAL", "VALOR", "VALOR BDI", "VALOR COM BDI"]
                ):
                    continue

                # Verificar coluna E para totais
                cell_coef_check = sheet.cell(row=y, column=col_coef_idx).value
                if cell_coef_check:
                    coef_upper = str(cell_coef_check).upper()
                    if any(x in coef_upper for x in totals_pular):
                        continue
                    if any(
                        x in coef_upper
                        for x in ["TOTAL", "VALOR", "VALOR BDI", "VALOR COM BDI"]
                    ):
                        continue
                    if any(x in coef_upper for x in valores_pular):
                        continue

                # Filtro por código
                codigo_item = desc.strip()
                codigo_upper = codigo_item.upper()

                if inicia_por and not codigo_upper.startswith(inicia_por.upper()):
                    continue
                if nao_inicia_por and codigo_upper.startswith(nao_inicia_por.upper()):
                    continue

                # Adicionar fórmula
                adicionar_formula = verificar_codigo_valido(sheet, y, col_desc_idx)
                if not adicionar_formula:
                    continue

                if fator_coef:
                    cell_coef = sheet.cell(row=y, column=col_coef_idx)
                    if not isinstance(cell_coef, MergedCell):
                        if not (
                            cell_coef.value
                            and isinstance(cell_coef.value, str)
                            and cell_coef.value.startswith("=")
                        ):
                            formula = (
                                f"={get_column_letter(col_coef_antigo_idx)}{y}*FATOR"
                            )
                            cell_coef.value = formula
                            linhas_adicionadas.append(f"L{y}: {desc[:30]} -> {formula}")
                else:
                    cell_preco = sheet.cell(row=y, column=col_preco_idx)
                    if not isinstance(cell_preco, MergedCell):
                        if not (
                            cell_preco.value
                            and isinstance(cell_preco.value, str)
                            and cell_preco.value.startswith("=")
                        ):
                            formula = f"=ROUND({get_column_letter(col_preco_antigo_idx)}{y}*FATOR, 2)"
                            cell_preco.value = formula
                            linhas_adicionadas.append(f"L{y}: {desc[:30]} -> {formula}")

    if linhas_adicionadas:
        print(f"   Adicionadas em Composições: {len(linhas_adicionadas)}")
        for item in linhas_adicionadas[:10]:
            print(f"      {item}")
        if len(linhas_adicionadas) > 10:
            print(f"      ... e mais {len(linhas_adicionadas) - 10}")

    return len(linhas_adicionadas)


def verificar_fator_aux(
    sheet,
    col_desc_idx,
    col_coef_idx,
    col_preco_idx,
    col_coef_antigo_idx,
    col_preco_antigo_idx,
    itens_fator,
    totals_pular,
    valores_pular,
):
    """
    Verifica e adiciona fórmulas de fator na planilha auxiliar.
    """
    print(">>> Verificando e adicionando fatores (Auxiliares)...")

    linhas_adicionadas = []
    max_row = sheet.max_row

    # Encontrar todas as seções
    secoes_encontradas = {}
    nomes_processados = set()

    for item_info in itens_fator:
        nome = item_info["nome"]
        total_str = item_info["total"]
        nome_upper = nome.upper()
        total_upper = total_str.upper() if total_str else ""

        cache_key = nome_upper
        if cache_key in nomes_processados:
            continue

        nomes_processados.add(cache_key)

        secoes = encontrar_todas_secoes(
            sheet, col_desc_idx, nome_upper, total_upper, "", ""
        )
        if secoes:
            secoes_encontradas[cache_key] = secoes
            print(f"   Encontradas {len(secoes)} seções de '{nome_upper}'")

    # Processar cada item
    for item_info in itens_fator:
        nome = item_info["nome"]
        fator_coef = item_info["fatorCoeficiente"]
        inicia_por = item_info["iniciaPor"]
        nao_inicia_por = item_info["naoIniciaPor"]

        nome_upper = nome.upper()
        cache_key = nome_upper

        if cache_key not in secoes_encontradas:
            print(f"   !! Não encontrou nenhuma seção para: {nome_upper}")
            continue

        secoes = secoes_encontradas[cache_key]

        for idx, (inicio, fim) in enumerate(secoes):
            for y in range(inicio + 1, fim):
                cell_desc = sheet.cell(row=y, column=col_desc_idx).value
                if isinstance(cell_desc, MergedCell):
                    continue

                desc = str(cell_desc) if cell_desc else ""
                desc_upper = desc.upper()

                # Pular linhas de títulos
                if "COEFICIENTE" in desc_upper or "PREÇO UNITÁRIO" in desc_upper:
                    continue

                cell_fonte = sheet.cell(row=y, column=col_desc_idx + 2).value
                cell_unid = sheet.cell(row=y, column=col_desc_idx + 3).value
                if cell_fonte and "FONTE" in str(cell_fonte).upper():
                    continue
                if cell_unid and "UNID" in str(cell_unid).upper():
                    continue

                # Pular totais
                if any(x in desc_upper for x in totals_pular):
                    continue
                if any(
                    x in desc_upper
                    for x in ["TOTAL", "VALOR", "VALOR BDI", "VALOR COM BDI"]
                ):
                    continue

                # Verificar coluna E
                cell_coef_check = sheet.cell(row=y, column=col_coef_idx).value
                if cell_coef_check:
                    coef_upper = str(cell_coef_check).upper()
                    if any(x in coef_upper for x in totals_pular):
                        continue
                    if any(
                        x in coef_upper
                        for x in ["TOTAL", "VALOR", "VALOR BDI", "VALOR COM BDI"]
                    ):
                        continue
                    if any(x in coef_upper for x in valores_pular):
                        continue

                # Filtro por código
                codigo_item = desc.strip()
                codigo_upper = codigo_item.upper()

                if inicia_por and not codigo_upper.startswith(inicia_por.upper()):
                    continue
                if nao_inicia_por and codigo_upper.startswith(nao_inicia_por.upper()):
                    continue

                # Adicionar fórmula
                adicionar_formula = verificar_codigo_valido(sheet, y, col_desc_idx)
                if not adicionar_formula:
                    continue

                if fator_coef:
                    cell_coef = sheet.cell(row=y, column=col_coef_idx)
                    if not isinstance(cell_coef, MergedCell):
                        if not (
                            cell_coef.value
                            and isinstance(cell_coef.value, str)
                            and cell_coef.value.startswith("=")
                        ):
                            formula = (
                                f"={get_column_letter(col_coef_antigo_idx)}{y}*FATOR"
                            )
                            cell_coef.value = formula
                            linhas_adicionadas.append(f"L{y}: {desc[:30]} -> {formula}")
                else:
                    cell_preco = sheet.cell(row=y, column=col_preco_idx)
                    if not isinstance(cell_preco, MergedCell):
                        if not (
                            cell_preco.value
                            and isinstance(cell_preco.value, str)
                            and cell_preco.value.startswith("=")
                        ):
                            formula = f"=ROUND({get_column_letter(col_preco_antigo_idx)}{y}*FATOR, 2)"
                            cell_preco.value = formula
                            linhas_adicionadas.append(f"L{y}: {desc[:30]} -> {formula}")

    if linhas_adicionadas:
        print(f"   Adicionadas em Auxiliares: {len(linhas_adicionadas)}")
        for item in linhas_adicionadas[:10]:
            print(f"      {item}")
        if len(linhas_adicionadas) > 10:
            print(f"      ... e mais {len(linhas_adicionadas) - 10}")

    return len(linhas_adicionadas)


def verificar_codigo_valido(sheet, row, col_desc_idx):
    """
    Verifica se a célula contém um código válido de item.
    Códigos válidos começam com dígito OU letra seguida de números (ex: I00378S, S10555, I00081)
    Textos descritivos como "Material", "Mão de Obra" não devem receber fórmula.
    """
    cell_item_val = sheet.cell(row=row, column=col_desc_idx).value
    if not cell_item_val:
        return False

    cell_item_str = str(cell_item_val).strip()
    if not cell_item_str:
        return False

    is_digit_start = cell_item_str[0].isdigit()
    is_alpha_with_digits = cell_item_str[0].isalpha() and any(
        c.isdigit() for c in cell_item_str
    )

    return is_digit_start or is_alpha_with_digits


def encontrar_todas_secoes(
    sheet, col_desc_idx, nome_upper, total_upper, inicia_por="", nao_inicia_por=""
):
    """
    Encontra TODAS as seções com o nome dado na planilha.
    Retorna lista de tuplas (inicio, fim) para cada seção.
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
            cell_upper = cell_str.upper()
            if cell and nome_upper in cell_upper:
                if "TOTAL" not in cell_upper:
                    if inicia_por:
                        if not cell_str.startswith(inicia_por):
                            continue
                    else:
                        if not cell_upper.startswith(nome_upper):
                            continue
                        if len(cell_upper) > len(nome_upper):
                            next_char = cell_upper[len(nome_upper)]
                            if next_char not in [":"]:
                                continue
                    if nao_inicia_por and cell_upper.startswith(nao_inicia_por.upper()):
                        continue
                    inicios.append((i, col_busca))
                    break

    # Para cada início, encontrar o fim mais próximo
    for inicio, col_inicio in inicios:
        fim = -1
        for j in range(inicio + 1, max_row + 1):
            for col_busca in colunas_busca:
                cell = sheet.cell(row=j, column=col_busca).value
                cell_upper = str(cell).upper() if cell else ""
                if cell and cell_upper.endswith(total_upper):
                    fim = j
                    break
            if fim != -1:
                break
        if fim != -1:
            secoes.append((inicio, fim))

    return secoes
