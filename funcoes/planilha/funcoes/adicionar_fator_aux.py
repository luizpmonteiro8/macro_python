from openpyxl.utils import column_index_from_string
from funcoes.common.buscar_palavras import (
    buscar_palavra_com_linha,
    buscar_palavra_com_linha_exato,
    buscar_palavra_com_linha_iniciando,
)
from funcoes.common.copiar_coluna import copiar_coluna_com_numeros
from funcoes.common.valor_bdi_final import valor_bdi_final
from funcoes.get.get_linhas_json import *


def copiar_colunas(sheet, dados):
    copiar_coluna_com_numeros(
        sheet, get_coeficiente_aux(dados), get_copiar_coeficiente_aux(dados)
    )
    copiar_coluna_com_numeros(
        sheet, get_preco_unitario_aux(dados), get_copiar_preco_unitario_aux(dados)
    )


def adicionar_formula_preco_unitario_menos_preco_antigo(sheet, dados):
    origem = get_copiar_preco_unitario_aux(dados)
    destino_idx = column_index_from_string(origem) + 1
    preco_unit = get_preco_unitario_aux(dados)

    for i, cell in enumerate(sheet[origem], start=1):
        if cell.value is not None:
            sheet.cell(row=i, column=destino_idx).value = (
                f"=({origem}{i}-{preco_unit}{i})"
            )


def fator_nos_item_totais_aux(
    sheet,
    dados,
    lin_ini,
    lin_fim,
    nome,
    totalNome,
    coeficiente,
    adicionar_fator,
    inicia_por=None,
    nao_inicia_por=None,
):
    col_desc = get_descricao_aux(dados)
    col_totais = get_coluna_totais_aux(dados)
    col_valor = get_valor_totais_aux(dados)
    col_preco = get_preco_unitario_aux(dados)
    col_coef = get_coeficiente_aux(dados)
    col_preco_antigo = get_copiar_preco_unitario_aux(dados)
    col_coef_antigo = get_copiar_coeficiente_aux(dados)

    inicial = buscar_palavra_com_linha_exato(sheet, col_desc, nome, lin_ini, lin_fim)
    final = buscar_palavra_com_linha_exato(
        sheet, col_totais, totalNome, lin_ini, lin_fim
    )

    if -1 < inicial < final:
        sheet[f"{col_valor}{final}"].value = (
            f"=SUM({col_valor}{inicial+1}:{col_valor}{final-1})"
        )

        for y in range(inicial + 1, final):
            desc = sheet[f"{col_desc}{y}"].value
            if inicia_por and (desc is None or not desc.startswith(inicia_por)):
                continue
            if nao_inicia_por and (desc is None or desc.startswith(nao_inicia_por)):
                continue

            if coeficiente and adicionar_fator:
                sheet[f"{col_coef}{y}"].value = f"={col_coef_antigo}{y}*FATOR"
            elif adicionar_fator:
                sheet[f"{col_preco}{y}"].value = (
                    f"=ROUND({col_preco_antigo}{y}*FATOR, 2)"
                )

        return inicial, final


# no topo do módulo
cache_encontrados = {}
cache_nao_encontrados = {}


def buscar_auxiliar_no_aux(workbook, dados, itemChave, lin, lin_total, nivel=1):
    if nivel > 50:
        print("⚠️ Recursão profunda demais — possível loop infinito.")
        return

    sheet_aux = workbook[get_planilha_aux(dados)]
    ultima_linha = sheet_aux.max_row

    col_item = get_item_descricao_comp_aux(dados)
    col_desc = get_descricao_aux(dados)
    col_valor = get_valor_totais_aux(dados)
    col_preco = get_preco_unitario_aux(dados)
    col_totais = get_coluna_totais_aux(dados)
    val_str = get_valor_string(dados)

    itens_array = [v for k, v in itemChave.items() if k.startswith("item")]

    for x in range(lin, lin_total):
        cod = sheet_aux[f"{col_desc}{x}"].value
        item = sheet_aux[f"{col_item}{x}"].value
        if item is None:
            continue

        chave_busca = f"{cod} {item}"

        # pula item que já sabemos que não existe
        if chave_busca in cache_nao_encontrados:
            continue

        # usa valor já encontrado
        if chave_busca in cache_encontrados:
            linha_ini = cache_encontrados[chave_busca]
        else:
            print(f"busca item na auxiliar: {chave_busca} na linha {x}")
            linha_ini = buscar_palavra_com_linha(
                sheet_aux, col_desc, chave_busca, 1, ultima_linha
            )
            if linha_ini == -1:
                linha_ini = buscar_palavra_com_linha(
                    sheet_aux, col_desc, f"{cod} ", 1, ultima_linha
                )
            if linha_ini == -1:
                linha_ini = buscar_palavra_com_linha_iniciando(
                    sheet_aux, col_desc, chave_busca, 1, ultima_linha
                )

            # atualiza cache
            if linha_ini == -1:
                cache_nao_encontrados[chave_busca] = True
                print(f"⚠️ Item não encontrado na auxiliar: {chave_busca}")
                continue
            else:
                cache_encontrados[chave_busca] = linha_ini
                print(
                    f"✅ Item encontrado na auxiliar: {chave_busca} — linha {linha_ini}"
                )

        linha_fim = buscar_palavra_com_linha_exato(
            sheet_aux, col_totais, val_str, linha_ini, ultima_linha
        )
        if linha_fim <= 0:
            cache_nao_encontrados[chave_busca] = True
            print(f"⚠️ Item não encontrado na auxiliar (valor final): {chave_busca}")
            continue

        if not (linha_ini <= x <= linha_fim):
            sheet_aux[f"{col_preco}{x}"].value = (
                f"='COMPOSICOES AUXILIARES'!{col_valor}{linha_fim}"
            )

        final_total_linha_array = set()
        for item_cfg in itens_array:
            resultado_fator = fator_nos_item_totais_aux(
                sheet_aux,
                dados,
                linha_ini,
                linha_fim,
                item_cfg["nome"],
                item_cfg["total"],
                item_cfg["fatorCoeficiente"] == "Sim",
                item_cfg["adicionarFator"] == "Sim",
                item_cfg.get("iniciaPor"),
                item_cfg.get("naoIniciaPor"),
            )

            if resultado_fator:
                linha_desc, linha_total = resultado_fator
                final_total_linha_array.add(linha_total)

                if (
                    item_cfg.get("buscarAuxiliar") == "Sim"
                    and linha_desc > 0
                    and linha_total > 0
                ):
                    buscar_auxiliar_no_aux(
                        workbook, dados, itemChave, linha_desc, linha_total, nivel + 1
                    )

        if final_total_linha_array:
            linha_valor_sum = buscar_palavra_com_linha(
                sheet_aux, col_totais, val_str, linha_ini, linha_fim + 1
            )
            if linha_valor_sum > 0:
                sheet_aux[f"{col_valor}{linha_valor_sum}"].value = (
                    f"=SUM({','.join(f'{col_valor}{linha}' for linha in final_total_linha_array)})"
                )


def adicionar_fator_totais_aux(workbook, dados, itemChave, lin_ini, lin_fim):
    sheet_aux = workbook[get_planilha_aux(dados)]
    col_totais = get_coluna_totais_aux(dados)
    col_valor = get_valor_totais_aux(dados)
    val_str = get_valor_string(dados)

    itens_array = [v for k, v in itemChave.items() if k.startswith("item")]

    final_total_linha_array = set()

    for item_cfg in itens_array:
        resultado_fator = fator_nos_item_totais_aux(
            sheet_aux,
            dados,
            lin_ini,
            lin_fim,
            item_cfg["nome"],
            item_cfg["total"],
            item_cfg["fatorCoeficiente"] == "Sim",
            item_cfg["adicionarFator"] == "Sim",
            item_cfg.get("iniciaPor"),
            item_cfg.get("naoIniciaPor"),
        )

        if resultado_fator:
            linha_desc, linha_total = resultado_fator
            final_total_linha_array.add(linha_total)

            if (
                item_cfg.get("buscarAuxiliar") == "Sim"
                and linha_desc > 0
                and linha_total > 0
            ):
                buscar_auxiliar_no_aux(
                    workbook, dados, itemChave, linha_desc, linha_total
                )

    if final_total_linha_array:
        linha_valor_sum = buscar_palavra_com_linha(
            sheet_aux, col_totais, val_str, lin_ini, lin_fim + 1
        )
        if linha_valor_sum > 0:
            sheet_aux[f"{col_valor}{linha_valor_sum}"].value = (
                f"=SUM({','.join(f'{col_valor}{linha}' for linha in final_total_linha_array)})"
            )


def adicionar_fator_aux(workbook, dados):
    sheet_aux = workbook[get_planilha_aux(dados)]
    copiar_colunas(sheet_aux, dados)
    adicionar_formula_preco_unitario_menos_preco_antigo(sheet_aux, dados)
    valor_bdi_final(
        sheet_aux, dados, get_coluna_totais_aux(dados), get_valor_totais_aux(dados)
    )
