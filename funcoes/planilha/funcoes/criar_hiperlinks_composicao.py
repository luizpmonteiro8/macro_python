from funcoes.common.buscar_palavras import (
    buscar_palavra_contem,
    buscar_palavra_com_linha,
)
from funcoes.get.get_linhas_json import (
    get_planilha_orcamentaria,
    get_planilha_comp,
    get_planilha_codigo,
    get_planilha_descricao,
    get_planilha_preco_unitario,
    get_descricao_comp,
    get_coluna_totais_comp,
    get_valor_totais_comp,
    get_valor_com_bdi_string,
)
from funcoes.planilha.funcoes.adicionar_fator_comp import (
    criar_link_composicao_orcamentaria,
)


def criar_hiperlinks_composicao(workbook, dados, lin_ini, lin_fim):
    # usado no planilha orcamentaria
    sheet = workbook[get_planilha_orcamentaria(dados)]
    sheet_comp = workbook[get_planilha_comp(dados)]
    sheet_comp_max = sheet_comp.max_row + 1

    col_preco_planilha = get_planilha_preco_unitario(dados)
    col_desc_comp = get_descricao_comp(dados)
    col_totais_comp = get_coluna_totais_comp(dados)
    col_valor_comp = get_valor_totais_comp(dados)
    col_cod = get_planilha_codigo(dados)
    col_desc = get_planilha_descricao(dados)
    valor_com_bdi = get_valor_com_bdi_string(dados)

    linha_busca_ini = 1
    itens_nao_encontrados = []

    for x in range(lin_ini, lin_fim):
        cod = sheet[f"{col_cod}{x}"].value
        descricao = sheet[f"{col_desc}{x}"].value
        if descricao is None:
            continue

        print(f"busca item {cod} {descricao} na linha {x}")

        linha_ini_comp = -1

        # 1ª tentativa: busca por "código + descrição" usando contém
        if linha_ini_comp == -1:
            linha_ini_comp = buscar_palavra_contem(
                sheet_comp,
                col_desc_comp,
                f"{cod} {descricao}",
                linha_busca_ini,
                sheet_comp_max,
            )

        # 2ª tentativa: busca por "código + descrição" no início da linha
        if linha_ini_comp == -1:
            linha_ini_comp = buscar_palavra_com_linha(
                sheet_comp,
                col_desc_comp,
                f"{cod} {descricao}",
                linha_busca_ini,
                sheet_comp_max,
            )

        # 3ª tentativa: busca pelo código apenas
        if linha_ini_comp == -1:
            linha_ini_comp = buscar_palavra_contem(
                sheet_comp,
                col_desc_comp,
                cod,
                linha_busca_ini,
                sheet_comp_max,
            )

        # 4ª tentativa: busca pelo código do início da planilha
        if linha_ini_comp == -1:
            linha_ini_comp = buscar_palavra_contem(
                sheet_comp,
                col_desc_comp,
                cod,
                1,
                sheet_comp_max,
            )

        if linha_ini_comp == -1:
            itens_nao_encontrados.append(f"{cod} {descricao}")
            print(f"⚠️ Item não encontrado na composição: {cod} {descricao}")
            continue

        print(f"encontrado item {cod} {descricao} -> linha do título: {linha_ini_comp}")

        # Busca a linha com "VALOR COM BDI" após o título da composição
        linha_valor_bdi = buscar_palavra_com_linha(
            sheet_comp,
            col_totais_comp,
            valor_com_bdi,
            linha_ini_comp,
            sheet_comp_max,
        )

        if linha_valor_bdi <= 0:
            itens_nao_encontrados.append(
                f"{cod} {descricao} (VALOR COM BDI não encontrado)"
            )
            print(f"⚠️ VALOR COM BDI não encontrado para: {cod} {descricao}")
            continue

        print(f">>> VALOR COM BDI encontrado na linha: {linha_valor_bdi}")

        # Cria hyperlink na descrição (aponta para o título da composição)
        criar_link_composicao_orcamentaria(
            sheet_origem=sheet,
            col_desc=col_desc,
            linha_origem=x,
            nome_planilha_destino=get_planilha_comp(dados),
            col_valor=col_valor_comp,
            linha_destino=linha_ini_comp,
        )

        # Cria fórmula na coluna PREÇO UNITÁRIO R$ (aponta para VALOR COM BDI)
        sheet[f"{col_preco_planilha}{x}"].value = (
            f"={get_planilha_comp(dados)}!{col_valor_comp}{linha_valor_bdi}"
        )

        linha_busca_ini = max(1, linha_valor_bdi)

    if itens_nao_encontrados:
        print(f"\n⚠️ Itens não encontrados ({len(itens_nao_encontrados)}):")
        for item in itens_nao_encontrados:
            print(f"  - {item}")
        return (
            False,
            f"ERRO: {len(itens_nao_encontrados)} itens não encontrados na composição",
        )

    return True, None
