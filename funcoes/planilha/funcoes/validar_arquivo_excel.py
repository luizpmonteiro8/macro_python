import openpyxl

from funcoes.common.buscar_palavras import buscar_palavra


def validar_arquivo_excel(workbook, dados):
    """
    Valida a estrutura do arquivo Excel antes do processamento.

    Verifica:
    1. Planilha orçamentária existe e tem os cabeçalhos corretos
    2. Planilha RESUMO existe e tem valores correspondentes
    3. Planilha COMPOSICOES existe e tem valores correspondentes
    4. Planilha COMPOSICOES AUXILIARES existe e tem valores correspondentes

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON

    Returns:
        tuple: (True, None) se válido, ou (False, mensagem_erro) se inválido
    """

    # ============================================
    # CABEÇALHOS E VALORES A VERIFICAR
    # ============================================

    # Cabeçalhos esperados na planilha orçamentária
    cabeçalhos_esperados = [
        "ITEM",
        "CÓDIGO",
        "DESCRIÇÃO",
        "UND",
        "QUANTIDADE",
    ]

    # Valores a verificar na planilha orçamentária
    valores_orcamentaria = [
        dados.get("valorFinal", "VALOR BDI TOTAL"),
    ]

    # Valores a verificar na planilha RESUMO
    valores_resumo = [
        dados.get("valorTotalResumo", "VALOR TOTAL RESUMO:"),
    ]

    # Valores a verificar na planilha COMPOSICOES
    valores_composicao = [
        dados.get("valor", "VALOR:"),
        dados.get("valorBdi", "VALOR BDI:"),
        dados.get("valorTotal", "VALOR TOTAL:"),
        dados.get("valorComBdi", "VALOR COM BDI:"),
    ]

    # Valores a verificar na planilha COMPOSICOES AUXILIARES (mesmos que COMPOSICOES)
    valores_auxiliar = [
        dados.get("valor", "VALOR:"),
        dados.get("valorBdi", "VALOR BDI:"),
        dados.get("valorTotal", "VALOR TOTAL:"),
        dados.get("valorComBdi", "VALOR COM BDI:"),
    ]

    # ============================================
    # OBTÉM NOMES DAS PLANILHAS DO JSON
    # ============================================

    nome_planilha_orcamentaria = dados.get(
        "planilhaOrcamentaria", "PLANILHA ORCAMENTARIA"
    )
    nome_planilha_fator = dados.get("planilhaFator", "RESUMO")
    nome_planilha_composicao = dados.get("planilhaComposicao", "COMPOSICOES")
    nome_planilha_auxiliar = dados.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")

    # Coluna e valor inicial para buscar a linha dos cabeçalhos
    coluna_inicial = dados.get("colunaInicial", "A")
    valor_inicial = dados.get("valorInicial", "ITEM")

    # ============================================
    # Verificação 1: Planilha Orçamentária existe
    # ============================================
    print(f">>> Verificando planilha orçamentária: '{nome_planilha_orcamentaria}'")

    if nome_planilha_orcamentaria not in workbook.sheetnames:
        mensagem = (
            f"ERRO: Planilha '{nome_planilha_orcamentaria}' não encontrada!\n\n"
            f"Planilhas disponíveis: {', '.join(workbook.sheetnames)}"
        )
        return False, mensagem

    sheet_orcamentaria = workbook[nome_planilha_orcamentaria]

    # ============================================
    # Verificação 2: Linha de cabeçalhos existe
    # Usa buscar_palavra para achar a linha com "ITEM" na coluna configurada
    # ============================================
    print(f">>> Buscando linha com '{valor_inicial}' na coluna '{coluna_inicial}'...")

    linha_cabecalhos = buscar_palavra(sheet_orcamentaria, coluna_inicial, valor_inicial)

    if linha_cabecalhos == -1:
        mensagem = (
            f"ERRO: Valor '{valor_inicial}' não encontrado na coluna '{coluna_inicial}'!\n\n"
            f"Não foi possível encontrar a linha de cabeçalhos na planilha."
        )
        return False, mensagem

    print(f">>> Linha de cabeçalhos encontrada na linha {linha_cabecalhos + 1}")

    # Obter valores da linha de cabeçalhos
    valores_linha = []
    for cell in sheet_orcamentaria[linha_cabecalhos + 1]:
        if cell.value is not None:
            valores_linha.append(str(cell.value).strip().upper())
        else:
            valores_linha.append("")

    print(f">>> Cabeçalhos encontrados: {valores_linha}")

    # Verificar cabeçalhos obrigatórios
    cabeçalhos_faltantes = []
    for cabecalho in cabeçalhos_esperados:
        encontrado = False
        for valor in valores_linha:
            if cabecalho.upper() in valor or valor in cabecalho.upper():
                encontrado = True
                break
        if not encontrado:
            cabeçalhos_faltantes.append(cabecalho)

    if cabeçalhos_faltantes:
        mensagem = (
            f"ERRO: Cabeçalhos não encontrados na linha {linha_cabecalhos + 1}!\n\n"
            f"Cabeçalhos esperados: {', '.join(cabeçalhos_esperados)}\n"
            f"Cabeçalhos encontrados: {', '.join(valores_linha)}\n"
            f"Faltando: {', '.join(cabeçalhos_faltantes)}"
        )
        return False, mensagem

    print(f">>> Todos os cabeçalhos obrigatórios encontrados!")

    # Verificar valores na planilha orçamentária
    print(f">>> Verificando valores na planilha orçamentária...")
    valores_faltantes_orcamentaria = []

    for valor_buscado in valores_orcamentaria:
        if not valor_buscado:
            continue

        valor_encontrado = False
        for row in sheet_orcamentaria.iter_rows():
            for cell in row:
                if cell.value and valor_buscado.upper() in str(cell.value).upper():
                    valor_encontrado = True
                    print(
                        f">>> Valor '{valor_buscado}' encontrado na célula {cell.coordinate}"
                    )
                    break
            if valor_encontrado:
                break

        if not valor_encontrado:
            valores_faltantes_orcamentaria.append(valor_buscado)

    if valores_faltantes_orcamentaria:
        mensagem = (
            f"ERRO: Valores não encontrados na planilha '{nome_planilha_orcamentaria}'!\n\n"
            f"Valores faltando: {', '.join(valores_faltantes_orcamentaria)}"
        )
        return False, mensagem

    # ============================================
    # Verificação 3: Planilha RESUMO existe
    # ============================================
    print(f">>> Verificando planilha RESUMO: '{nome_planilha_fator}'")

    if nome_planilha_fator not in workbook.sheetnames:
        mensagem = (
            f"ERRO: Planilha '{nome_planilha_fator}' não encontrada!\n\n"
            f"Planilhas disponíveis: {', '.join(workbook.sheetnames)}"
        )
        return False, mensagem

    sheet_resumo = workbook[nome_planilha_fator]

    # Verificar valores na RESUMO
    print(f">>> Verificando valores na planilha RESUMO...")
    valores_faltantes_resumo = []

    for valor_buscado in valores_resumo:
        if not valor_buscado:
            continue

        valor_encontrado = False
        for row in sheet_resumo.iter_rows():
            for cell in row:
                if cell.value and valor_buscado.upper() in str(cell.value).upper():
                    valor_encontrado = True
                    print(
                        f">>> Valor '{valor_buscado}' encontrado na célula {cell.coordinate}"
                    )
                    break
            if valor_encontrado:
                break

        if not valor_encontrado:
            valores_faltantes_resumo.append(valor_buscado)

    if valores_faltantes_resumo:
        mensagem = (
            f"ERRO: Valores não encontrados na planilha '{nome_planilha_fator}'!\n\n"
            f"Valores faltando: {', '.join(valores_faltantes_resumo)}"
        )
        return False, mensagem

    # ============================================
    # Verificação 4: Planilha COMPOSICOES existe
    # ============================================
    print(f">>> Verificando planilha COMPOSICOES: '{nome_planilha_composicao}'")

    if nome_planilha_composicao not in workbook.sheetnames:
        mensagem = (
            f"ERRO: Planilha '{nome_planilha_composicao}' não encontrada!\n\n"
            f"Planilhas disponíveis: {', '.join(workbook.sheetnames)}"
        )
        return False, mensagem

    sheet_composicao = workbook[nome_planilha_composicao]

    # Verificar valores na COMPOSICOES
    print(f">>> Verificando valores na planilha COMPOSICOES...")
    valores_faltantes = []

    for valor_buscado in valores_composicao:
        if not valor_buscado:
            continue

        valor_encontrado = False
        for row in sheet_composicao.iter_rows():
            for cell in row:
                if cell.value and valor_buscado.upper() in str(cell.value).upper():
                    valor_encontrado = True
                    print(
                        f">>> Valor '{valor_buscado}' encontrado na célula {cell.coordinate}"
                    )
                    break
            if valor_encontrado:
                break

        if not valor_encontrado:
            valores_faltantes.append(valor_buscado)

    if valores_faltantes:
        mensagem = (
            f"ERRO: Valores não encontrados na planilha '{nome_planilha_composicao}'!\n\n"
            f"Valores faltando: {', '.join(valores_faltantes)}"
        )
        return False, mensagem

    # ============================================
    # Verificação 5: Planilha COMPOSICOES AUXILIARES existe
    # ============================================
    print(
        f">>> Verificando planilha COMPOSICOES AUXILIARES: '{nome_planilha_auxiliar}'"
    )

    if nome_planilha_auxiliar not in workbook.sheetnames:
        mensagem = (
            f"ERRO: Planilha '{nome_planilha_auxiliar}' não encontrada!\n\n"
            f"Planilhas disponíveis: {', '.join(workbook.sheetnames)}"
        )
        return False, mensagem

    sheet_auxiliar = workbook[nome_planilha_auxiliar]

    # Verificar valores na COMPOSICOES AUXILIARES
    print(f">>> Verificando valores na planilha COMPOSICOES AUXILIARES...")

    for valor_buscado in valores_auxiliar:
        if not valor_buscado:
            continue

        valor_encontrado = False
        for row in sheet_auxiliar.iter_rows():
            for cell in row:
                if cell.value and valor_buscado.upper() in str(cell.value).upper():
                    valor_encontrado = True
                    print(
                        f">>> Valor '{valor_buscado}' encontrado na célula {cell.coordinate}"
                    )
                    break
            if valor_encontrado:
                break

        if not valor_encontrado:
            valores_faltantes.append(f"{valor_buscado} (em COMPOSICOES AUXILIARES)")

    if valores_faltantes:
        mensagem = (
            f"ERRO: Valores não encontrados nas planilhas!\n\n"
            f"Valores faltando: {', '.join(valores_faltantes)}"
        )
        return False, mensagem

    # ============================================
    # Todas as verificações passaram
    # ============================================
    print(f">>> [OK] Validação do arquivo concluída com sucesso!")
    return True, None
