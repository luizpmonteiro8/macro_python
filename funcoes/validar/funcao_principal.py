"""
Função principal para validar configurações e arquivo Excel.
Orquestra todos os componentes de validação.
"""

import json

import openpyxl
from tkinter import messagebox

from funcoes.validar.validar_estrutura_base import validar_estrutura_base
from funcoes.validar.validar_planilha_orcamentaria import validar_planilha_orcamentaria
from funcoes.validar.validar_planilha_resumo import validar_planilha_resumo
from funcoes.validar.validar_planilha_composicoes import validar_planilha_composicoes
from funcoes.validar.validar_planilha_composicoes_auxiliares import (
    validar_planilha_composicoes_auxiliares,
)


CAMINHO_JSON = "config/valores_colunas.json"


def validar_arquivo_excel(filepath, dados):
    """
    Valida a estrutura completa do arquivo Excel antes do processamento.

    Se encontrar erros que podem ser corrigidos, exibe janela para correção
    e salva automaticamente no arquivo JSON. Recarrega o workbook e os dados
    após cada correção para garantir que as mudanças sejam refletidas.

    Args:
        filepath: Caminho do arquivo Excel
        dados: Lista de configurações do arquivo JSON

    Returns:
        tuple: (True, workbook, dados_atualizados) se válido,
               ou (False, None, None) se inválido/cancelado
    """
    if not isinstance(dados, list):
        dados = [dados]

    indice_config = 0

    print("=" * 60)
    print(">>> INICIANDO VALIDAÇÃO DO ARQUIVO EXCEL")
    print("=" * 60)

    while True:
        erros = []
        workbook = openpyxl.load_workbook(filepath)

        dados_atualizados = json.load(open(CAMINHO_JSON, "r", encoding="utf-8"))
        json_antes = json.dumps(dados_atualizados, sort_keys=True)

        print("\n>>> [FASE 1] Validando estrutura base...")
        if not validar_estrutura_base(
            workbook, dados_atualizados[indice_config], erros
        ):
            mensagem = "ERROS NA VALIDAÇÃO:\n" + "\n".join(erros)
            return False, None, None

        json_depois = json.dumps(
            json.load(open(CAMINHO_JSON, "r", encoding="utf-8")), sort_keys=True
        )
        if json_antes != json_depois:
            print(">>> [INFO] Configurações atualizadas. Recarregando...")
            workbook.close()
            continue
        json_antes = json_depois

        print(">>> [OK] Estrutura base válida")

        print("\n>>> [FASE 2] Validando planilha orçamentária...")
        valido, sheet_orcamentaria, linha_cabecalhos = validar_planilha_orcamentaria(
            workbook, dados_atualizados, erros, indice_config
        )

        if not valido:
            print(
                ">>> [ERRO] Validação da planilha orçamentária falhou ou foi cancelada."
            )
            print(">>> Encerrando validação.")
            workbook.close()
            return False, None, None

        json_depois = json.dumps(
            json.load(open(CAMINHO_JSON, "r", encoding="utf-8")), sort_keys=True
        )
        if json_antes != json_depois:
            print(">>> [INFO] Nome da planilha orçamentária corrigido. Recarregando...")
            workbook.close()
            continue
        json_antes = json_depois

        print(
            f">>> [OK] Planilha orçamentária válida (cabeçalhos na linha {linha_cabecalhos + 1 if linha_cabecalhos else '?'})"
        )

        print("\n>>> [FASE 3] Validando planilha RESUMO...")
        valido, sheet_resumo = validar_planilha_resumo(
            workbook, dados_atualizados, erros, indice_config
        )

        if not valido:
            print(">>> [ERRO] Validação da planilha RESUMO falhou ou foi cancelada.")
            print(">>> Encerrando validação.")
            workbook.close()
            return False, None, None

        json_depois = json.dumps(
            json.load(open(CAMINHO_JSON, "r", encoding="utf-8")), sort_keys=True
        )
        if json_antes != json_depois:
            print(">>> [INFO] Nome da planilha RESUMO corrigido. Recarregando...")
            workbook.close()
            continue
        json_antes = json_depois

        print("\n>>> [FASE 4] Validando planilha COMPOSIÇÕES...")
        valido, sheet_composicao = validar_planilha_composicoes(
            workbook, dados_atualizados, erros, indice_config
        )

        if not valido:
            print(
                ">>> [ERRO] Validação da planilha COMPOSIÇÕES falhou ou foi cancelada."
            )
            print(">>> Encerrando validação.")
            workbook.close()
            return False, None, None

        json_depois = json.dumps(
            json.load(open(CAMINHO_JSON, "r", encoding="utf-8")), sort_keys=True
        )
        if json_antes != json_depois:
            print(">>> [INFO] Nome da planilha COMPOSIÇÕES corrigido. Recarregando...")
            workbook.close()
            continue
        json_antes = json_depois

        print("\n>>> [FASE 5] Validando planilha COMPOSIÇÕES AUXILIARES...")
        valido, sheet_auxiliar = validar_planilha_composicoes_auxiliares(
            workbook, dados_atualizados, erros, indice_config
        )

        if not valido:
            print(
                ">>> [ERRO] Validação da planilha AUXILIARES falhou ou foi cancelada."
            )
            print(">>> Encerrando validação.")
            workbook.close()
            return False, None, None

        json_depois = json.dumps(
            json.load(open(CAMINHO_JSON, "r", encoding="utf-8")), sort_keys=True
        )
        if json_antes != json_depois:
            print(">>> [INFO] Nome da planilha AUXILIARES corrigido. Recarregando...")
            workbook.close()
            continue

        print("\n" + "=" * 60)
        print(">>> [OK] VALIDAÇÃO CONCLUÍDA COM SUCESSO!")
        print("=" * 60)
        return True, workbook, dados_atualizados


def validar_todas_configuracoes(caminho_arquivo):
    """
    Valida todas as configurações do arquivo Excel e JSON.
    Se usuário clicar Cancelar em qualquer etapa, interrompe o processo.

    Args:
        caminho_arquivo: Caminho do arquivo Excel a validar

    Returns:
        bool: True se todas as validações passaram, False caso contrário
    """
    erros = []

    if not openpyxl.load_workbook(caminho_arquivo):
        erros.append(f"ERRO: O arquivo '{caminho_arquivo}' não foi encontrado.")
        messagebox.showerror("Erro", "\n".join(erros))
        return False

    try:
        with open(CAMINHO_JSON, "r", encoding="utf-8") as f:
            dados = json.load(f)
    except Exception as e:
        erros.append(f"ERRO: Não foi possível carregar as configurações: {str(e)}")
        messagebox.showerror("Erro", "\n".join(erros))
        return False

    valido, workbook, dados_atualizados = validar_arquivo_excel(
        caminho_arquivo, dados
    )

    if workbook:
        workbook.close()

    return valido
