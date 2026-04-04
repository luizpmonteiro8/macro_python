import openpyxl
from openpyxl.utils import column_index_from_string

from funcoes.common.buscar_palavras import buscar_palavra
from funcoes.get.get_linhas_json import *


# ============================================
# FUNÇÕES AUXILIARES DE VALIDAÇÃO
# ============================================


def validar_estrutura_base(workbook, dados, erros):
    """
    Valida a estrutura básica: workbook e dados JSON.

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros (será adicionada mensagens de erro)

    Returns:
        bool: True se válido, False se inválido
    """
    if workbook is None:
        erros.append("ERRO: O arquivo Excel não pôde ser aberto. Verifique se o arquivo existe e não está corrompido.")
        return False

    if not hasattr(workbook, "sheetnames"):
        erros.append("ERRO: O arquivo Excel está em um formato inválido ou corrompido.")
        return False

    if dados is None:
        erros.append("ERRO: As configurações do sistema não foram carregadas. Entre em contato com o suporte.")
        return False

    if not isinstance(dados, dict):
        erros.append("ERRO: As configurações do sistema estão em um formato inesperado. Entre em contato com o suporte.")
        return False

    return True


def validar_nome_planilha(nome_planilha, nome_exibicao, erros):
    """
    Valida que o nome da planilha não está vazio.

    Args:
        nome_planilha: Nome da planilha a validar
        nome_exibicao: Nome para exibição nas mensagens de erro
        erros: Lista de erros

    Returns:
        bool: True se válido, False se inválido
    """
    if nome_planilha is None or nome_planilha.strip() == "":
        erros.append(f"ERRO: O nome da aba '{nome_exibicao}' não foi definido nas configurações.")
        return False
    return True


def validar_planilha_existe(workbook, nome_planilha, nome_exibicao, erros):
    """
    Valida que uma planilha existe no workbook.

    Args:
        workbook: Objeto workbook do openpyxl
        nome_planilha: Nome da planilha a verificar
        nome_exibicao: Nome para exibição nas mensagens de erro
        erros: Lista de erros

    Returns:
        tuple: (existe: bool, sheet: Worksheet ou None)
    """
    if nome_planilha not in workbook.sheetnames:
        erros.append(
            f"ERRO: A aba '{nome_exibicao}' não foi encontrada no arquivo Excel.\n"
            f"Abas disponíveis no arquivo: {', '.join(workbook.sheetnames)}"
        )
        return False, None
    return True, workbook[nome_planilha]


def validar_coluna_existe(sheet, nome_coluna, nome_exibicao, erros):
    """
    Valida que uma coluna é válida (letra de coluna válida).

    Args:
        sheet: Worksheet do openpyxl
        nome_coluna: Letra da coluna (ex: 'A', 'B', 'AA')
        nome_exibicao: Nome para exibição nas mensagens de erro
        erros: Lista de erros

    Returns:
        bool: True se válido, False se inválido
    """
    if not nome_coluna or nome_coluna.strip() == "":
        erros.append(f"ERRO: A coluna '{nome_exibicao}' não foi definida nas configurações.")
        return False

    try:
        column_index_from_string(nome_coluna)
    except Exception:
        erros.append(
            f"ERRO: A coluna '{nome_coluna}' ({nome_exibicao}) não existe na planilha."
        )
        return False

    return True


def validar_celula_bdi(sheet, coluna_fator, linha_fator, erros):
    """
    Valida que a célula do BDI existe e tem um valor numérico.

    Args:
        sheet: Worksheet do openpyxl
        coluna_fator: Letra da coluna do BDI (ex: 'G')
        linha_fator: Número da linha do BDI (ex: 4)
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, valor_bdi: float ou None)
    """
    if not coluna_fator or not linha_fator:
        erros.append("ERRO: A localização do BDI (taxa de benefícios) não foi configurada corretamente.")
        return False, None

    try:
        linha = int(linha_fator)
        if linha <= 0:
            erros.append(
                f"ERRO: O número da linha do BDI deve ser maior que zero."
            )
            return False, None
    except (ValueError, TypeError):
        erros.append(f"ERRO: O valor '{linha_fator}' não é um número válido para a linha do BDI.")
        return False, None

    cell = sheet[f"{coluna_fator}{linha}"]
    if cell.value is None:
        erros.append(f"ERRO: A célula de BDI (coluna {coluna_fator}, linha {linha}) está vazia.")
        return False, None

    try:
        valor_bdi = float(str(cell.value).replace(",", "."))
        return True, valor_bdi
    except (ValueError, TypeError):
        erros.append(
            f"ERRO: O valor do BDI '{cell.value}' não é um número válido. O BDI deve ser um número (exemplo: 28,55)."
        )
        return False, None


def validar_valor_existe_na_coluna(
    sheet, coluna, valor_buscado, nome_valor, nome_planilha, erros
):
    """
    Valida que um valor existe em uma coluna específica.

    Args:
        sheet: Worksheet do openpyxl
        coluna: Letra da coluna para buscar
        valor_buscado: Valor a buscar
        nome_valor: Nome do valor para exibição
        nome_planilha: Nome da planilha para mensagens de erro
        erros: Lista de erros

    Returns:
        bool: True se encontrado, False se não encontrado
    """
    if not valor_buscado or valor_buscado.strip() == "":
        return True

    linha_encontrada = buscar_palavra(sheet, coluna, valor_buscado)

    if linha_encontrada == -1:
        erros.append(
            f"ERRO: O texto '{valor_buscado}' ({nome_valor}) não foi encontrado na coluna '{coluna}' da aba '{nome_planilha}'."
        )
        return False

    return True


def validar_valores_existem_na_coluna(
    sheet, coluna, valores_dict, nome_planilha, erros
):
    """
    Valida que múltiplos valores existem em uma coluna.

    Args:
        sheet: Worksheet do openpyxl
        coluna: Letra da coluna para buscar
        valores_dict: Dicionário {nome: valor} dos valores a buscar
        nome_planilha: Nome da planilha para mensagens de erro
        erros: Lista de erros

    Returns:
        int: Número de erros encontrados
    """
    erros_iniciais = len(erros)

    for nome_valor, valor_buscado in valores_dict.items():
        if not valor_buscado or valor_buscado.strip() == "":
            continue
        validar_valor_existe_na_coluna(
            sheet, coluna, valor_buscado, nome_valor, nome_planilha, erros
        )

    return len(erros) - erros_iniciais


# ============================================
# FUNÇÕES DE VALIDAÇÃO POR PLANILHA
# ============================================


def validar_planilha_orcamentaria(workbook, dados, erros):
    """
    Valida a planilha orçamentária.

    Valida:
    - Nome da planilha existe
    - Planilha não está vazia
    - Coluna inicial existe
    - Coluna final existe (para busca de valor total)
    - Valor inicial (ITEM) existe na coluna inicial
    - Cabeçalhos esperados existem
    - Valor final existe na coluna inicial

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None, linha_cabecalhos: int ou None)
    """
    nome_planilha = dados.get("planilhaOrcamentaria", "PLANILHA ORCAMENTARIA")

    if not validar_nome_planilha(nome_planilha, "Orçamentária", erros):
        return False, None, None

    existe, sheet = validar_planilha_existe(
        workbook, nome_planilha, "Orçamentária", erros
    )
    if not existe:
        return False, None, None

    if sheet.max_row < 2:
        erros.append(
            f"ERRO: A aba '{nome_planilha}' está vazia ou não tem dados suficientes."
        )
        return False, sheet, None

    coluna_inicial = dados.get("colunaInicial", "A")
    valor_inicial = dados.get("valorInicial", "ITEM")
    valor_final = dados.get("valorFinal", "VALOR BDI TOTAL")

    if not validar_coluna_existe(sheet, coluna_inicial, "Inicial", erros):
        return False, sheet, None

    coluna_final = dados.get("colunaFinal", "F")
    if not validar_coluna_existe(sheet, coluna_final, "Coluna Final", erros):
        return False, sheet, None

    linha_cabecalhos = buscar_palavra(sheet, coluna_inicial, valor_inicial)

    if linha_cabecalhos == -1:
        erros.append(
            f"ERRO: O texto '{valor_inicial}' não foi encontrado na coluna '{coluna_inicial}'.\n"
            f"Isso pode significar que a estrutura da aba '{nome_planilha}' está diferente do esperado."
        )
        return False, sheet, None

    valores_linha = []
    for cell in sheet[linha_cabecalhos + 1]:
        if cell.value is not None:
            valores_linha.append(str(cell.value).strip().upper())
        else:
            valores_linha.append("")

    cabecalhos_esperados = ["ITEM", "CÓDIGO", "DESCRIÇÃO", "UND", "QUANTIDADE"]

    cabecalhos_faltantes = []
    for cabecalho in cabecalhos_esperados:
        encontrado = False
        for valor in valores_linha:
            if cabecalho.upper() in valor or valor in cabecalho.upper():
                encontrado = True
                break
        if not encontrado:
            cabecalhos_faltantes.append(cabecalho)

    if cabecalhos_faltantes:
        erros.append(
            f"ERRO: Algumas colunas obrigatórias não foram encontradas na linha {linha_cabecalhos + 1} da aba '{nome_planilha}'.\n"
            f"Colunas esperadas: {', '.join(cabecalhos_esperados)}\n"
            f"Colunas encontradas: {', '.join(valores_linha)}\n"
            f"Faltando: {', '.join(cabeçalhos_faltantes)}"
        )
        return False, sheet, linha_cabecalhos

    colunas_opcionais = [
        ("planilhaCodigo", "Código"),
        ("planilhaDescricao", "Descrição"),
        ("planilhaQuantidade", "Quantidade"),
        ("planilhaPrecoUnitario", "Preço Unitário"),
        ("planilhaPrecoTotal", "Preço Total"),
        ("planilhaPrecoUnitarioCopiar", "Preço Unitário Copiar"),
    ]

    for key, nome in colunas_opcionais:
        col = dados.get(key)
        if col:
            validar_coluna_existe(sheet, col, nome, erros)

    if valor_final:
        validar_valor_existe_na_coluna(
            sheet, coluna_inicial, valor_final, "Valor Final", nome_planilha, erros
        )

    return (
        (len(erros) == 0 or all("ERRO" not in e for e in erros)),
        sheet,
        linha_cabecalhos,
    )


def validar_planilha_resumo(workbook, dados, erros):
    """
    Valida a planilha RESUMO (FATOR).

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None)
    """
    nome_planilha = dados.get("planilhaFator", "RESUMO")

    if not validar_nome_planilha(nome_planilha, "Resumo", erros):
        return False, None

    existe, sheet = validar_planilha_existe(workbook, nome_planilha, "Resumo", erros)
    if not existe:
        return False, None

    if sheet.max_row < 2:
        erros.append(
            f"ERRO: A aba '{nome_planilha}' está vazia ou não tem dados suficientes."
        )
        return False, sheet

    coluna_fator = dados.get("colunaFator", "G")
    linha_fator = dados.get("linhaFator", "4")

    if not validar_coluna_existe(sheet, coluna_fator, "Fator", erros):
        return False, sheet

    bdi_valido, _ = validar_celula_bdi(sheet, coluna_fator, linha_fator, erros)

    valor_total_resumo = dados.get("valorTotalResumo", "VALOR TOTAL RESUMO:")

    if valor_total_resumo:
        validar_valor_existe_na_coluna(
            sheet,
            coluna_fator,
            valor_total_resumo,
            "Valor Total do Resumo",
            nome_planilha,
            erros,
        )

    return True, sheet


def validar_planilha_composicoes(workbook, dados, erros):
    """
    Valida a planilha COMPOSIÇÕES.

    Valida:
    - Nome da planilha existe
    - Planilha não está vazia
    - Todas as colunas configuradas existem
    - Coluna de totais (colunaTotaisComposicao) existe
    - Coluna de valor totais (valorTotaisComposicao) existe
    - Valores de busca (valorComBdi, valorBdi, valorTotal, valor) existem na coluna de totais

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None)
    """
    nome_planilha = dados.get("planilhaComposicao", "COMPOSICOES")

    if not validar_nome_planilha(nome_planilha, "Composições", erros):
        return False, None

    existe, sheet = validar_planilha_existe(
        workbook, nome_planilha, "Composições", erros
    )
    if not existe:
        return False, None

    if sheet.max_row < 2:
        erros.append(
            f"ERRO: A aba '{nome_planilha}' está vazia ou não tem dados suficientes."
        )
        return False, sheet

    colunas_comp = [
        ("composicaoDescricao", "Descrição"),
        ("colunaItemDescricaoComposicao", "Código do Item"),
        ("composicaoCoeficiente", "Coeficiente"),
        ("composicaoPrecoUnitario", "Preço Unitário"),
        ("composicaoCoeficienteCopiar", "Coeficiente (copiar)"),
        ("composicaoPrecoUnitarioCopiar", "Preço Unitário (copiar)"),
        ("colunaTotaisComposicao", "Coluna de Totais"),
    ]

    for key, nome in colunas_comp:
        col = dados.get(key)
        if col:
            if not validar_coluna_existe(sheet, col, nome, erros):
                erros.append(
                    f"ERRO: A coluna '{col}' ({nome}) não foi encontrada na aba '{nome_planilha}'."
                )

    col_totais = get_coluna_totais_comp(dados)
    if not col_totais:
        erros.append(f"ERRO: A coluna de totais da aba Composições não foi configurada.")
        return False, sheet

    if not validar_coluna_existe(sheet, col_totais, "Coluna de Totais", erros):
        return False, sheet

    col_valor_totais = get_valor_totais_comp(dados)
    if not validar_coluna_existe(sheet, col_valor_totais, "Coluna de Valores Totais", erros):
        return False, sheet

    valores_a_verificar = {
        "valor_com_bdi": get_valor_com_bdi_string(dados),
        "valor_bdi": get_valor_bdi_comp(dados),
        "valor_total": get_valor_total_string(dados),
        "valor_string": get_valor_string(dados),
    }

    erros_iniciais = len(erros)
    validar_valores_existem_na_coluna(
        sheet, col_totais, valores_a_verificar, nome_planilha, erros
    )

    return len(erros) == erros_iniciais, sheet


def validar_planilha_composicoes_auxiliares(workbook, dados, erros):
    """
    Valida a planilha COMPOSIÇÕES AUXILIARES.

    Valida:
    - Nome da planilha existe
    - Planilha não está vazia
    - Todas as colunas configuradas existem
    - Coluna de totais (colunaTotaisAuxiliar) existe
    - Coluna de valor totais (valorTotaisAuxiliar) existe
    - Valores de busca (valorComBdi, valorBdi, valorTotal, valor) existem na coluna de totais

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None)
    """
    nome_planilha = dados.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")

    if not validar_nome_planilha(nome_planilha, "Auxiliares", erros):
        return False, None

    existe, sheet = validar_planilha_existe(
        workbook, nome_planilha, "Auxiliares", erros
    )
    if not existe:
        return False, None

    if sheet.max_row < 2:
        erros.append(
            f"ERRO: A aba '{nome_planilha}' está vazia ou não tem dados suficientes."
        )
        return False, sheet

    colunas_aux = [
        ("auxiliarDescricao", "Descrição"),
        ("auxiliarCoeficiente", "Coeficiente"),
        ("auxiliarPrecoUnitario", "Preço Unitário"),
        ("auxiliarCoeficienteCopiar", "Coeficiente (copiar)"),
        ("auxiliarPrecoUnitarioCopiar", "Preço Unitário (copiar)"),
        ("colunaTotaisAuxiliar", "Coluna de Totais"),
    ]

    for key, nome in colunas_aux:
        col = dados.get(key)
        if col:
            if not validar_coluna_existe(sheet, col, nome, erros):
                erros.append(
                    f"ERRO: A coluna '{col}' ({nome}) não foi encontrada na aba '{nome_planilha}'."
                )

    col_totais = get_coluna_totais_aux(dados)
    if not col_totais:
        erros.append(f"ERRO: A coluna de totais da aba Auxiliares não foi configurada.")
        return False, sheet

    if not validar_coluna_existe(sheet, col_totais, "Coluna de Totais", erros):
        return False, sheet

    col_valor_totais_aux = get_valor_totais_aux(dados)
    if not validar_coluna_existe(sheet, col_valor_totais_aux, "Coluna de Valores Totais", erros):
        return False, sheet

    valores_a_verificar = {
        "valor_com_bdi": get_valor_com_bdi_string(dados),
        "valor_bdi": get_valor_bdi_comp(dados),
        "valor_total": get_valor_total_string(dados),
        "valor_string": get_valor_string(dados),
    }

    erros_iniciais = len(erros)
    validar_valores_existem_na_coluna(
        sheet, col_totais, valores_a_verificar, nome_planilha, erros
    )

    return len(erros) == erros_iniciais, sheet


# ============================================
# FUNÇÃO PRINCIPAL DE VALIDAÇÃO
# ============================================


def validar_arquivo_excel(workbook, dados):
    """
    Valida a estrutura completa do arquivo Excel antes do processamento.

    Verifica:
    1. Estrutura base (workbook e dados JSON)
    2. Planilha orçamentária existe e tem os cabeçalhos corretos
       - Coluna inicial e final existem
       - Valor inicial (ITEM) e valor final existem
    3. Planilha RESUMO existe e tem valores correspondentes
       - Célula BDI configurada corretamente
       - Valor total resumo existe
    4. Planilha COMPOSIÇÕES existe e tem valores correspondentes
       - Todas as colunas configuradas existem
       - Coluna de totais e coluna de valor totais existem
       - Valores de busca existem na coluna de totais
    5. Planilha COMPOSIÇÕES AUXILIARES existe e tem valores correspondentes
       - Todas as colunas configuradas existem
       - Coluna de totais e coluna de valor totais existem
       - Valores de busca existem na coluna de totais

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON

    Returns:
        tuple: (True, None) se válido, ou (False, mensagem_erro) se inválido
    """
    erros = []

    print("=" * 60)
    print(">>> INICIANDO VALIDAÇÃO DO ARQUIVO EXCEL")
    print("=" * 60)

    # ============================================
    # FASE 1: Validação de Estrutura Base
    # ============================================
    print("\n>>> [FASE 1] Validando estrutura base...")
    if not validar_estrutura_base(workbook, dados, erros):
        mensagem = "ERROS NA VALIDAÇÃO:\n" + "\n".join(erros)
        return False, mensagem

    print(">>> [OK] Estrutura base válida")

    # ============================================
    # FASE 2: Validar Planilha Orçamentária
    # ============================================
    print("\n>>> [FASE 2] Validando planilha orçamentária...")
    valido, sheet_orcamentaria, linha_cabecalhos = validar_planilha_orcamentaria(
        workbook, dados, erros
    )

    if not valido:
        print(">>> [ERRO] Problemas encontrados na planilha orçamentária")
    else:
        print(
            f">>> [OK] Planilha orçamentária válida (cabeçalhos na linha {linha_cabecalhos + 1})"
        )

    # ============================================
    # FASE 3: Validar Planilha RESUMO
    # ============================================
    print("\n>>> [FASE 3] Validando planilha RESUMO...")
    valido, sheet_resumo = validar_planilha_resumo(workbook, dados, erros)

    if not valido:
        print(">>> [ERRO] Problemas encontrados na planilha RESUMO")
    else:
        print(">>> [OK] Planilha RESUMO válida")

    # ============================================
    # FASE 4: Validar Planilha COMPOSIÇÕES
    # ============================================
    print("\n>>> [FASE 4] Validando planilha COMPOSIÇÕES...")
    valido, sheet_composicao = validar_planilha_composicoes(workbook, dados, erros)

    if not valido:
        print(">>> [ERRO] Problemas encontrados na planilha COMPOSIÇÕES")
    else:
        print(">>> [OK] Planilha COMPOSIÇÕES válida")

    # ============================================
    # FASE 5: Validar Planilha COMPOSIÇÕES AUXILIARES
    # ============================================
    print("\n>>> [FASE 5] Validando planilha COMPOSIÇÕES AUXILIARES...")
    valido, sheet_auxiliar = validar_planilha_composicoes_auxiliares(
        workbook, dados, erros
    )

    if not valido:
        print(">>> [ERRO] Problemas encontrados na planilha AUXILIARES")
    else:
        print(">>> [OK] Planilha AUXILIARES válida")

    # ============================================
    # RESULTADO FINAL
    # ============================================
    print("\n" + "=" * 60)

    if erros:
        print(f">>> [ERRO] Validação encontrou {len(erros)} problema(s):")
        for i, erro in enumerate(erros, 1):
            print(f"    {i}. {erro}")

        mensagem = "ERROS NA VALIDAÇÃO:\n" + "\n".join(
            [f"{i+1}. {e}" for i, e in enumerate(erros)]
        )
        return False, mensagem

    print(">>> [OK] VALIDAÇÃO CONCLUÍDA COM SUCESSO!")
    print("=" * 60)
    return True, None


# ============================================
# FUNÇÃO LEGACY (para compatibilidade)
# ============================================


def validar_arquivo_excel_legacy(workbook, dados):
    """
    Versão original da função de validação (mantida para compatibilidade).
    Use validar_arquivo_excel() para nova implementação.
    """
    pass
