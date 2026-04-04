"""
Componente para validar a estrutura base do arquivo Excel e dados JSON.
"""

import openpyxl


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
        erros.append(
            "ERRO: O arquivo Excel não pôde ser aberto. Verifique se o arquivo existe e não está corrompido."
        )
        return False

    if not hasattr(workbook, "sheetnames"):
        erros.append("ERRO: O arquivo Excel está em um formato inválido ou corrompido.")
        return False

    if dados is None:
        erros.append(
            "ERRO: As configurações do sistema não foram carregadas. Entre em contato com o suporte."
        )
        return False

    if not isinstance(dados, dict):
        erros.append(
            "ERRO: As configurações do sistema estão em um formato inesperado. Entre em contato com o suporte."
        )
        return False

    return True
