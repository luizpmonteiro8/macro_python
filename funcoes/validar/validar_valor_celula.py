"""
Componente para validar se um valor existe em uma célula do Excel.
Inclui tratamento correto do botão Cancelar.
"""

from openpyxl.utils import column_index_from_string

from funcoes.validar.janela_corrigir import janela_corrigir_valor
from funcoes.validar.salvar_json import salvar_json_corrigido


def validar_valor_celula(
    sheet,
    nome_coluna,
    nome_linha,
    valor_esperado,
    nome_exibicao,
    erros,
    dados,
    indice_config,
    nome_planilha,
    campo_json,
):
    """
    Valida que um valor existe na célula especificada.
    TRATAMENTO CORRETO: Se usuário cancelar, retorna False imediatamente.

    Args:
        sheet: Worksheet do openpyxl
        nome_coluna: Letra da coluna (ex: 'A', 'B', 'AA')
        nome_linha: Número da linha (ex: 1, 2, 3)
        valor_esperado: Valor esperado na célula
        nome_exibicao: Nome para exibição nas mensagens de erro
        erros: Lista de erros
        dados: Dados do JSON (para correção)
        indice_config: Índice da configuração
        nome_planilha: Nome da aba para instrução
        campo_json: Nome do campo no JSON para correção

    Returns:
        bool: True se válido, False se inválido
    """
    if not nome_coluna or nome_coluna.strip() == "":
        # LOOP: continuar pedindo até acertar ou cancelar
        while True:
            instrucao = (
                f"1. Abra o arquivo Excel\n"
                f"2. Vá até a aba '{nome_planilha}'\n"
                f"3. Observe as letras no topo das colunas (A, B, C, D...)\n"
                f"4. Digite a letra da coluna que contém '{nome_exibicao}'"
            )
            confirmado, novo_valor = janela_corrigir_valor(
                titulo="Coluna não definida",
                mensagem=f"A coluna para '{nome_exibicao}' não foi definida nas configurações.",
                instrucao=instrucao,
                valor_atual="",
                valor_default="A",
            )

            # Se cancelou, para
            if not confirmado:
                return False

            # Se digitou algo
            if novo_valor:
                try:
                    column_index_from_string(novo_valor)
                    dados[indice_config][campo_json] = novo_valor.upper()
                    salvar_json_corrigido(dados, indice_config)
                    return True
                except Exception:
                    from tkinter import messagebox

                    messagebox.showerror(
                        "Erro",
                        f"'{novo_valor}' não é uma coluna válida! Tente novamente.",
                    )
                    continue  # Continua o loop

        return False

    if not nome_linha or str(nome_linha).strip() == "":
        # LOOP: continuar pedindo até acertar ou cancelar
        while True:
            instrucao = (
                f"1. Abra o arquivo Excel\n"
                f"2. Vá até a aba '{nome_planilha}'\n"
                f"3. Observe os números à esquerda das linhas (1, 2, 3...)\n"
                f"4. Digite o número da linha que contém '{nome_exibicao}'"
            )
            confirmado, novo_valor = janela_corrigir_valor(
                titulo="Linha não definida",
                mensagem=f"A linha para '{nome_exibicao}' não foi definida nas configurações.",
                instrucao=instrucao,
                valor_atual="",
                valor_default="1",
            )

            # Se cancelou, para
            if not confirmado:
                return False

            # Se digitou algo
            if novo_valor:
                try:
                    int(novo_valor)
                    dados[indice_config][campo_json] = str(novo_valor)
                    salvar_json_corrigido(dados, indice_config)
                    return True
                except Exception:
                    from tkinter import messagebox

                    messagebox.showerror(
                        "Erro",
                        f"'{novo_valor}' não é um número válido! Tente novamente.",
                    )
                    continue  # Continua o loop

        return False

    try:
        num_linha = int(nome_linha)
        num_coluna = column_index_from_string(nome_coluna)
        celula = sheet.cell(row=num_linha, column=num_coluna)
        valor_celula = str(celula.value).strip() if celula.value else ""
        valor_esperado_limpo = str(valor_esperado).strip() if valor_esperado else ""

        if valor_celula != valor_esperado_limpo:
            # LOOP: continuar pedindo até acertar ou cancelar
            while True:
                instrucao = (
                    f"1. Abra o arquivo Excel\n"
                    f"2. Vá até a aba '{nome_planilha}'\n"
                    f"3. Encontre a célula {nome_coluna}{nome_linha}\n"
                    f"4. O valor atual é: '{valor_celula}'\n"
                    f"5. Digite o valor correto que aparece nessa célula"
                )
                confirmado, novo_valor = janela_corrigir_valor(
                    titulo="Valor não encontrado",
                    mensagem=f"O valor '{valor_esperado}' não foi encontrado na célula {nome_coluna}{nome_linha}.",
                    instrucao=instrucao,
                    valor_atual=valor_esperado,
                    valor_default=valor_celula,
                )

                # Se cancelou, para
                if not confirmado:
                    return False

                # Se digitou algo
                if novo_valor:
                    dados[indice_config][campo_json] = novo_valor
                    salvar_json_corrigido(dados, indice_config)
                    return True

                # Se não digitou nada, volta a pedir
                continue

    except Exception as e:
        # LOOP: continuar pedindo até acertar ou cancelar
        while True:
            instrucao = (
                f"1. Abra o arquivo Excel\n"
                f"2. Vá até a aba '{nome_planilha}'\n"
                f"3. Verifique se a célula {nome_coluna}{nome_linha} existe\n"
                f"4. Digite o valor correto que aparece nessa célula"
            )
            confirmado, novo_valor = janela_corrigir_valor(
                titulo="Erro ao verificar valor",
                mensagem=f"Não foi possível verificar o valor '{valor_esperado}' na célula {nome_coluna}{nome_linha}: {str(e)}",
                instrucao=instrucao,
                valor_atual=valor_esperado,
                valor_default="",
            )

            # Se cancelou, para
            if not confirmado:
                return False

            # Se digitou algo
            if novo_valor:
                dados[indice_config][campo_json] = novo_valor
                salvar_json_corrigido(dados, indice_config)
                return True

            # Se não digitou nada, volta a pedir
            continue

    return True
