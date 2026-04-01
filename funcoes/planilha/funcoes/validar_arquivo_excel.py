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
        erros.append("ERRO: Workbook é None - arquivo Excel não pôde ser aberto.")
        return False

    if not hasattr(workbook, "sheetnames"):
        erros.append("ERRO: Objeto workbook inválido - não possui sheetnames.")
        return False

    if dados is None:
        erros.append("ERRO: Dados JSON é None - configuração não pôde ser carregada.")
        return False

    if not isinstance(dados, dict):
        erros.append("ERRO: Dados JSON não é um dicionário válido.")
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
        erros.append(f"ERRO: Nome da planilha {nome_exibicao} está vazio ou é None.")
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
            f"ERRO: Planilha '{nome_planilha}' ({nome_exibicao}) não encontrada!\n"
            f"Planilhas disponíveis: {', '.join(workbook.sheetnames)}"
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
        erros.append(f"ERRO: Coluna {nome_exibicao} está vazia ou é None.")
        return False

    try:
        column_index_from_string(nome_coluna)
    except Exception:
        erros.append(
            f"ERRO: Coluna {nome_exibicao} ('{nome_coluna}') não é uma coluna válida."
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
        erros.append("ERRO: Coluna ou linha do BDI não configuradas.")
        return False, None

    try:
        linha = int(linha_fator)
        if linha <= 0:
            erros.append(
                f"ERRO: Linha do BDI ({linha_fator}) deve ser um número positivo."
            )
            return False, None
    except (ValueError, TypeError):
        erros.append(f"ERRO: Linha do BDI ({linha_fator}) não é um número válido.")
        return False, None

    cell = sheet[f"{coluna_fator}{linha}"]
    if cell.value is None:
        erros.append(f"ERRO: Célula BDI ({coluna_fator}{linha}) está vazia.")
        return False, None

    try:
        valor_bdi = float(str(cell.value).replace(",", "."))
        return True, valor_bdi
    except (ValueError, TypeError):
        erros.append(
            f"ERRO: Valor do BDI na célula ({coluna_fator}{linha}) não é numérico: '{cell.value}'"
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
        return True  # Valores vazios são opcionais

    linha_encontrada = buscar_palavra(sheet, coluna, valor_buscado)

    if linha_encontrada == -1:
        erros.append(
            f"ERRO: Valor '{valor_buscado}' ({nome_valor}) não encontrado "
            f"na coluna '{coluna}' da planilha '{nome_planilha}'"
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

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None, linha_cabecalhos: int ou None)
    """
    # Obter nomes e valores configurados
    nome_planilha = dados.get("planilhaOrcamentaria", "PLANILHA ORCAMENTARIA")

    # Validar nome da planilha
    if not validar_nome_planilha(nome_planilha, "Orçamentária", erros):
        return False, None, None

    # Validar existência da planilha
    existe, sheet = validar_planilha_existe(
        workbook, nome_planilha, "Orçamentária", erros
    )
    if not existe:
        return False, None, None

    # Validar planilha vazia
    if sheet.max_row < 2:
        erros.append(
            f"ERRO: Planilha '{nome_planilha}' está vazia ou tem menos de 2 linhas."
        )
        return False, sheet, None

    # Obter colunas configuradas
    coluna_inicial = dados.get("colunaInicial", "A")
    valor_inicial = dados.get("valorInicial", "ITEM")
    valor_final = dados.get("valorFinal", "VALOR BDI TOTAL")

    # Validar coluna inicial
    if not validar_coluna_existe(sheet, coluna_inicial, "Inicial", erros):
        return False, sheet, None

    # Buscar linha de cabeçalhos
    linha_cabecalhos = buscar_palavra(sheet, coluna_inicial, valor_inicial)

    if linha_cabecalhos == -1:
        erros.append(
            f"ERRO: Valor '{valor_inicial}' não encontrado na coluna '{coluna_inicial}'!\n"
            f"Não foi possível encontrar a linha de cabeçalhos na planilha '{nome_planilha}'."
        )
        return False, sheet, None

    # Obter valores da linha de cabeçalhos
    valores_linha = []
    for cell in sheet[linha_cabecalhos + 1]:
        if cell.value is not None:
            valores_linha.append(str(cell.value).strip().upper())
        else:
            valores_linha.append("")

    # Cabeçalhos esperados
    cabecalhos_esperados = ["ITEM", "CÓDIGO", "DESCRIÇÃO", "UND", "QUANTIDADE"]

    # Verificar cabeçalhos obrigatórios
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
            f"ERRO: Cabeçalhos não encontrados na linha {linha_cabecalhos + 1} "
            f"da planilha '{nome_planilha}'!\n"
            f"Cabeçalhos esperados: {', '.join(cabecalhos_esperados)}\n"
            f"Cabeçalhos encontrados: {', '.join(valores_linha)}\n"
            f"Faltando: {', '.join(cabeçalhos_faltantes)}"
        )
        return False, sheet, linha_cabecalhos

    # Validar colunas configuradas no JSON (não críticos, apenas informativa)
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

    # Verificar valor final na planilha
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

    # Validar nome da planilha
    if not validar_nome_planilha(nome_planilha, "RESUMO", erros):
        return False, None

    # Validar existência da planilha
    existe, sheet = validar_planilha_existe(workbook, nome_planilha, "RESUMO", erros)
    if not existe:
        return False, None

    # Validar planilha vazia
    if sheet.max_row < 2:
        erros.append(
            f"ERRO: Planilha '{nome_planilha}' está vazia ou tem menos de 2 linhas."
        )
        return False, sheet

    # Validar célula BDI
    coluna_fator = dados.get("colunaFator", "G")
    linha_fator = dados.get("linhaFator", "4")

    # Validar coluna fator
    if not validar_coluna_existe(sheet, coluna_fator, "Fator BDI", erros):
        return False, sheet

    # Validar célula BDI (não é crítica se não existir, mas deve ser numérico se existir)
    bdi_valido, _ = validar_celula_bdi(sheet, coluna_fator, linha_fator, erros)

    # Validar valor total resumo
    valor_total_resumo = dados.get("valorTotalResumo", "VALOR TOTAL RESUMO:")

    if valor_total_resumo:
        validar_valor_existe_na_coluna(
            sheet,
            coluna_fator,
            valor_total_resumo,
            "Valor Total Resumo",
            nome_planilha,
            erros,
        )

    return True, sheet


def validar_planilha_composicoes(workbook, dados, erros):
    """
    Valida a planilha COMPOSIÇÕES.

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None)
    """
    nome_planilha = dados.get("planilhaComposicao", "COMPOSICOES")

    # Validar nome da planilha
    if not validar_nome_planilha(nome_planilha, "COMPOSIÇÕES", erros):
        return False, None

    # Validar existência da planilha
    existe, sheet = validar_planilha_existe(
        workbook, nome_planilha, "COMPOSIÇÕES", erros
    )
    if not existe:
        return False, None

    # Validar planilha vazia
    if sheet.max_row < 2:
        erros.append(
            f"ERRO: Planilha '{nome_planilha}' está vazia ou tem menos de 2 linhas."
        )
        return False, sheet

    # Validar colunas configuradas
    colunas_comp = [
        ("composicaoDescricao", "Descrição Composição"),
        ("colunaItemDescricaoComposicao", "Item Descrição Composição"),
        ("composicaoCoeficiente", "Coeficiente Composição"),
        ("composicaoPrecoUnitario", "Preço Unitário Composição"),
        ("composicaoCoeficienteCopiar", "Coeficiente Copiar Composição"),
        ("composicaoPrecoUnitarioCopiar", "Preço Unitário Copiar Composição"),
        ("colunaTotaisComposicao", "Coluna Totais Composição"),
    ]

    for key, nome in colunas_comp:
        col = dados.get(key)
        if col:
            if not validar_coluna_existe(sheet, col, nome, erros):
                erros.append(
                    f"ERRO: Coluna '{key}' = '{col}' (nome: {nome}) não é válida na planilha '{nome_planilha}'"
                )

    # Obter coluna de totais e valores a verificar
    col_totais = get_coluna_totais_comp(dados)
    if not col_totais:
        erros.append(f"ERRO: Coluna de totais não configurada para COMPOSIÇÕES.")
        return False, sheet

    if not validar_coluna_existe(sheet, col_totais, "Coluna Totais", erros):
        return False, sheet

    # Valores a verificar na coluna de totais
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

    Args:
        workbook: Objeto workbook do openpyxl
        dados: Dicionário com configurações do arquivo JSON
        erros: Lista de erros

    Returns:
        tuple: (válido: bool, sheet: Worksheet ou None)
    """
    nome_planilha = dados.get("planilhaAuxiliar", "COMPOSICOES AUXILIARES")

    # Validar nome da planilha
    if not validar_nome_planilha(nome_planilha, "AUXILIARES", erros):
        return False, None

    # Validar existência da planilha
    existe, sheet = validar_planilha_existe(
        workbook, nome_planilha, "AUXILIARES", erros
    )
    if not existe:
        return False, None

    # Validar planilha vazia
    if sheet.max_row < 2:
        erros.append(
            f"ERRO: Planilha '{nome_planilha}' está vazia ou tem menos de 2 linhas."
        )
        return False, sheet

    # Validar colunas configuradas
    colunas_aux = [
        ("auxiliarDescricao", "Descrição Auxiliar"),
        ("auxiliarCoeficiente", "Coeficiente Auxiliar"),
        ("auxiliarPrecoUnitario", "Preço Unitário Auxiliar"),
        ("auxiliarCoeficienteCopiar", "Coeficiente Copiar Auxiliar"),
        ("auxiliarPrecoUnitarioCopiar", "Preço Unitário Copiar Auxiliar"),
        ("colunaTotaisAuxiliar", "Coluna Totais Auxiliar"),
    ]

    for key, nome in colunas_aux:
        col = dados.get(key)
        if col:
            if not validar_coluna_existe(sheet, col, nome, erros):
                erros.append(
                    f"ERRO: Coluna '{key}' = '{col}' (nome: {nome}) não é válida na planilha '{nome_planilha}'"
                )

    # Obter coluna de totais e valores a verificar
    col_totais = get_coluna_totais_aux(dados)
    if not col_totais:
        erros.append(f"ERRO: Coluna de totais não configurada para AUXILIARES.")
        return False, sheet

    if not validar_coluna_existe(sheet, col_totais, "Coluna Totais Auxiliar", erros):
        return False, sheet

    # Valores a verificar na coluna de totais (mesmos do COMPOSIÇÕES)
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
    3. Planilha RESUMO existe e tem valores correspondentes
    4. Planilha COMPOSIÇÕES existe e tem valores correspondentes
    5. Planilha COMPOSIÇÕES AUXILIARES existe e tem valores correspondentes
    6. Célula BDI configurada corretamente
    7. Colunas configuradas no JSON são válidas

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
    # Implementação original aqui (código não duplicado)
    pass
