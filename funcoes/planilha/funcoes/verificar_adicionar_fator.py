from openpyxl.utils import column_index_from_string
from openpyxl.cell.cell import MergedCell


def verificar_adicionar_fator(workbook, dados):
    """
    Verifica se os fatores estão sendo adicionados corretamente.

    Lógica:
    - Se adicionarFator: "Sim" E fatorCoeficiente: "Sim" → usar coluna E (composicaoCoeficiente ou auxiliarCoeficiente)
    - Se adicionarFator: "Sim" E fatorCoeficiente: "Não" → usar coluna F (composicaoPrecoUnitario ou auxiliarPrecoUnitario)

    Valida em ambas as planilhas: composições e composições auxiliares.
    """
    print(">>> Verificando fatores dos itens...")

    # Obter nomes das planilhas do JSON
    planilha_comp = dados.get("planilhaComposicao", "COMPOSIÇÕES")
    planilha_aux = dados.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")

    # Obter colunas do JSON
    col_coef_comp = dados.get("composicaoCoeficiente", "E")
    col_preco_comp = dados.get("composicaoPrecoUnitario", "F")
    col_coef_aux = dados.get("auxiliarCoeficiente", "E")
    col_preco_aux = dados.get("auxiliarPrecoUnitario", "F")
    col_desc_comp = dados.get("composicaoDescricao", "A")
    col_desc_aux = dados.get("auxiliarDescricao", "A")

    # Construir mapa de itens que precisam de fator
    itens_com_fator = (
        {}
    )  # nome -> {"coeficiente": bool, "planilha": str, "coluna": str}

    if dados and isinstance(dados, list) and len(dados) > 0:
        primeiro_item = dados[0]
        for key, item in primeiro_item.items():
            if key.startswith("item") and isinstance(item, dict):
                if item.get("adicionarFator") == "Sim":
                    nome = item.get("nome", "")
                    fator_coef = item.get("fatorCoeficiente") == "Sim"
                    itens_com_fator[nome.upper()] = {
                        "fatorCoeficiente": fator_coef,
                        "total": item.get("total", ""),
                    }

    print(f">> Itens com adicionarFator: Sim - {len(itens_com_fator)}")
    for nome, info in itens_com_fator.items():
        print(f"   {nome}: fatorCoeficiente={info['fatorCoeficiente']}")

    # Verificar composições
    print(f"\n>> Verificando planilha: {planilha_comp}")
    itens_faltando_comp = verificar_planilha(
        workbook[planilha_comp],
        column_index_from_string(col_desc_comp),
        column_index_from_string(col_coef_comp),
        column_index_from_string(col_preco_comp),
        itens_com_fator,
    )

    # Verificar composições auxiliares
    print(f"\n>> Verificando planilha: {planilha_aux}")
    itens_faltando_aux = verificar_planilha(
        workbook[planilha_aux],
        column_index_from_string(col_desc_aux),
        column_index_from_string(col_coef_aux),
        column_index_from_string(col_preco_aux),
        itens_com_fator,
    )

    # Resumo
    total_faltando = len(itens_faltando_comp) + len(itens_faltando_aux)
    print(f"\n>>> Total de itens faltando fator: {total_faltando}")

    if itens_faltando_comp:
        print(f">> Em Composições ({len(itens_faltando_comp)}):")
        for item in itens_faltando_comp[:10]:
            print(f"   - {item}")
        if len(itens_faltando_comp) > 10:
            print(f"   ... e mais {len(itens_faltando_comp) - 10}")

    if itens_faltando_aux:
        print(f">> Em Composições Auxiliares ({len(itens_faltando_aux)}):")
        for item in itens_faltando_aux[:10]:
            print(f"   - {item}")
        if len(itens_faltando_aux) > 10:
            print(f"   ... e mais {len(itens_faltando_aux) - 10}")

    return total_faltando, itens_faltando_comp, itens_faltando_aux


def verificar_planilha(
    sheet, col_desc_idx, col_coef_idx, col_preco_idx, itens_com_fator
):
    """
    Verifica uma planilha específica para itens faltando fator.
    """
    itens_faltando = []
    max_row = sheet.max_row

    for i in range(1, max_row + 1):
        cell_desc = sheet.cell(row=i, column=col_desc_idx).value
        if not cell_desc or isinstance(cell_desc, MergedCell):
            continue

        descricao = str(cell_desc).strip().upper()

        # Verificar se a descrição contém algum dos itens que precisam de fator
        for nome_item, info in itens_com_fator.items():
            # Verificar tanto pelo nome quanto pela string "TOTAL xxx:"
            if nome_item in descricao or info["total"].upper() in descricao.upper():
                # Encontrou um item que precisa de fator
                if info["fatorCoeficiente"]:
                    # Deve ter valor na coluna E (coeficiente)
                    cell_coef = sheet.cell(row=i, column=col_coef_idx).value
                    if isinstance(cell_coef, MergedCell):
                        continue
                    if cell_coef is None or (
                        isinstance(cell_coef, str) and not cell_coef.startswith("=")
                    ):
                        itens_faltando.append(
                            f"{descricao[:50]} (coluna E/Coeficiente)"
                        )
                else:
                    # Deve ter valor na coluna F (preço unitário)
                    cell_preco = sheet.cell(row=i, column=col_preco_idx).value
                    if isinstance(cell_preco, MergedCell):
                        continue
                    if cell_preco is None or (
                        isinstance(cell_preco, str) and not cell_preco.startswith("=")
                    ):
                        itens_faltando.append(f"{descricao[:50]} (coluna F/Preço)")
                break

    return itens_faltando
