"""
Componente para validar se uma coluna existe no Excel.
Inclui tratamento correto do botão Cancelar.
"""

from openpyxl.utils import column_index_from_string

from funcoes.validar.janela_corrigir import janela_corrigir_valor
from funcoes.validar.salvar_json import salvar_json_corrigido


def validar_coluna_existe(
    sheet,
    nome_coluna,
    nome_exibicao,
    erros,
    dados,
    indice_config,
    nome_planilha,
    campo_json,
):
    """
    Valida que uma coluna é válida (letra de coluna válida).
    TRATAMENTO CORRETO: Se usuário cancelar, retorna False imediatamente.

    Args:
        sheet: Worksheet do openpyxl
        nome_coluna: Letra da coluna (ex: 'A', 'B', 'AA')
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
                mensagem=f"A coluna '{nome_exibicao}' não foi definida nas configurações.",
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

    try:
        column_index_from_string(nome_coluna)
    except Exception:
        # LOOP: continuar pedindo até acertar ou cancelar
        while True:
            instrucao = (
                f"1. Abra o arquivo Excel\n"
                f"2. Vá até a aba '{nome_planilha}'\n"
                f"3. Observe as letras no topo das colunas (A, B, C, D...)\n"
                f"4. Digite a letra da coluna correta para '{nome_exibicao}'"
            )
            confirmado, novo_valor = janela_corrigir_valor(
                titulo="Coluna inválida",
                mensagem=f"A coluna '{nome_coluna}' não é uma coluna válida do Excel.",
                instrucao=instrucao,
                valor_atual=nome_coluna,
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

    return True
