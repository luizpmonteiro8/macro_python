"""
Componente para validar a célula do BDI existe e tem um valor numérico.
"""

from tkinter import messagebox

from funcoes.validar.janela_corrigir import janela_corrigir_valor
from funcoes.validar.salvar_json import salvar_json_corrigido


CAMINHO_JSON = "config/valores_colunas.json"


def validar_celula_bdi(
    sheet,
    coluna_fator,
    linha_fator,
    erros,
    dados,
    indice_config,
    nome_planilha="RESUMO",
):
    """
    Valida que a célula do BDI existe e tem um valor numérico.

    Args:
        sheet: Worksheet do openpyxl
        coluna_fator: Letra da coluna do BDI (ex: 'G')
        linha_fator: Número da linha do BDI (ex: 4)
        erros: Lista de erros
        dados: Dados do JSON (para correção)
        indice_config: Índice da configuração
        nome_planilha: Nome da aba para instrução

    Returns:
        tuple: (válido: bool, valor_bdi: float ou None, corrigido: bool)
    """
    corrigido = False
    linha = 0

    if not coluna_fator or not linha_fator:
        instrucao = (
            f"1. Abra o arquivo Excel\n"
            f"2. Vá até a aba '{nome_planilha}'\n"
            f"3. Procure pela linha que contém 'BDI' ou 'TAXA BDI'\n"
            f"4. Digite o número da linha (geralmente 4 ou 5)"
        )
        confirmado, novo_valor = janela_corrigir_valor(
            titulo="Localização do BDI não definida",
            mensagem="A localização do BDI (taxa de benefícios) não foi configurada corretamente.",
            instrucao=instrucao,
            valor_atual=f"Coluna: {coluna_fator}, Linha: {linha_fator}",
            valor_default="4",
        )
        if confirmado and novo_valor:
            try:
                dados[indice_config]["linhaFator"] = str(int(novo_valor))
                salvar_json_corrigido(dados, indice_config)
                linha_fator = int(novo_valor)
                corrigido = True
            except ValueError:
                messagebox.showerror(
                    "Erro", f"O valor '{novo_valor}' não é um número válido!"
                )
                return False, None, False
        else:
            return False, None, False

    try:
        linha = int(linha_fator)
        if linha <= 0:
            instrucao = (
                f"1. Abra o arquivo Excel\n"
                f"2. Vá até a aba '{nome_planilha}'\n"
                f"3. Observe o número à esquerda da linha do BDI\n"
                f"4. Digite um número maior que zero"
            )
            confirmado, novo_valor = janela_corrigir_valor(
                titulo="Linha do BDI inválida",
                mensagem=f"O número da linha do BDI ({linha_fator}) deve ser maior que zero.",
                instrucao=instrucao,
                valor_atual=str(linha_fator),
                valor_default="4",
            )
            if confirmado and novo_valor:
                try:
                    dados[indice_config]["linhaFator"] = str(int(novo_valor))
                    salvar_json_corrigido(dados, indice_config)
                    linha_fator = int(novo_valor)
                    corrigido = True
                except ValueError:
                    return False, None, False
            else:
                return False, None, False
    except (ValueError, TypeError):
        instrucao = (
            f"1. Abra o arquivo Excel\n"
            f"2. Vá até a aba '{nome_planilha}'\n"
            f"3. Observe o número à esquerda da linha do BDI\n"
            f"4. Digite apenas números (exemplo: 4)"
        )
        confirmado, novo_valor = janela_corrigir_valor(
            titulo="Linha do BDI inválida",
            mensagem=f"O valor '{linha_fator}' não é um número válido para a linha do BDI.",
            instrucao=instrucao,
            valor_atual=str(linha_fator),
            valor_default="4",
        )
        if confirmado and novo_valor:
            try:
                dados[indice_config]["linhaFator"] = str(int(novo_valor))
                salvar_json_corrigido(dados, indice_config)
                linha_fator = int(novo_valor)
                corrigido = True
            except ValueError:
                messagebox.showerror(
                    "Erro", f"O valor '{novo_valor}' não é um número válido!"
                )
                return False, None, False
        else:
            return False, None, False

    cell = sheet[f"{coluna_fator}{linha}"]
    if cell.value is None:
        print(
            f">>> [AVISO] Célula BDI (coluna {coluna_fator}, linha {linha}) está vazia. Será preenchida automaticamente."
        )
        return True, None, False

    try:
        valor_bdi = float(str(cell.value).replace(",", "."))
        return True, valor_bdi, corrigido
    except (ValueError, TypeError):
        instrucao = (
            f"1. Abra o arquivo Excel\n"
            f"2. Vá até a aba '{nome_planilha}'\n"
            f"3. Procure pela célula do BDI\n"
            f"4. Digite o valor numérico do BDI (exemplo: 28,55)"
        )
        confirmado, novo_valor = janela_corrigir_valor(
            titulo="Valor do BDI inválido",
            mensagem=f"O valor do BDI '{cell.value}' não é um número válido. O BDI deve ser um número (exemplo: 28,55).",
            instrucao=instrucao,
            valor_atual=str(cell.value),
            valor_default="28,55",
        )
        if confirmado and novo_valor:
            try:
                valor_bdi = float(str(novo_valor).replace(",", "."))
                dados[indice_config]["BDI"] = str(valor_bdi).replace(".", ",")
                salvar_json_corrigido(dados, indice_config)
                return True, valor_bdi, True
            except ValueError:
                messagebox.showerror(
                    "Erro", f"O valor '{novo_valor}' não é um número válido!"
                )
                return False, None, False
        return False, None, False
