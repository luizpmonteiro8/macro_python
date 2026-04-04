"""
Componente para validar se um valor existe em uma coluna específica.
"""

from tkinter import messagebox

from funcoes.common.buscar_palavras import buscar_palavra
from funcoes.validar.janela_corrigir import janela_corrigir_valor
from funcoes.validar.salvar_json import salvar_json_corrigido


def validar_valor_existe_na_coluna(
    sheet,
    coluna,
    valor_buscado,
    nome_valor,
    nome_planilha,
    erros,
    dados,
    indice_config,
    campo_json,
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
        dados: Dados do JSON (para correção)
        indice_config: Índice da configuração
        campo_json: Nome do campo no JSON para correção

    Returns:
        bool: True se encontrado, False se não encontrado
    """
    if not valor_buscado or valor_buscado.strip() == "":
        return True

    linha_encontrada = buscar_palavra(sheet, coluna, valor_buscado)

    if linha_encontrada == -1:
        instrucao = (
            f"1. Abra o arquivo Excel\n"
            f"2. Vá até a aba '{nome_planilha}'\n"
            f"3. Procure na coluna '{coluna}' pelo texto relacionado a '{nome_valor}'\n"
            f"4. Digite o texto exatamente como aparece na célula"
        )
        while True:
            confirmado, novo_valor = janela_corrigir_valor(
                titulo="Texto não encontrado",
                mensagem=f"O texto '{valor_buscado}' não foi encontrado na coluna '{coluna}' da aba '{nome_planilha}'.",
                instrucao=instrucao,
                valor_atual=valor_buscado,
                valor_default=valor_buscado,
            )
            if not confirmado:
                return False
            if not novo_valor or novo_valor.strip() == "":
                messagebox.showerror(
                    "Erro", "O valor não pode ser vazio. Digite um valor válido."
                )
                continue
            linha_validada = buscar_palavra(sheet, coluna, novo_valor)
            if linha_validada == -1:
                messagebox.showerror(
                    "Erro",
                    f"O texto '{novo_valor}' não foi encontrado na coluna '{coluna}'.\n"
                    f"Verifique se o texto está correto e tente novamente.",
                )
                valor_buscado = novo_valor
                continue
            dados[indice_config][campo_json] = novo_valor
            salvar_json_corrigido(dados, indice_config)
            return True

    return True
