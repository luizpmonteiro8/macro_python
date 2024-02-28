import json
import tkinter as tk
from tkinter import messagebox

from funcoes.config.open_config import open_valores_colunas
from funcoes.view.interfaces.menu.funcoes.editar_coluna_nome import (
    atualizar_dropdown, local_abrir_janela_editar)

new_item = {
    "nome": "Novo",
    "planilhaOrcamentaria": "PLANILHA ORCAMENTARIA",
    "colunaInicial": "A",
    "colunaFinal": "F",
    "valorInicial": "ITEM",
    "valorFinal": "VALOR BDI TOTAL",
    "planilhaCodigo": "B",
    "planilhaDescricao": "C",
    "planilhaQuantidade": "F",
    "planilhaPrecoUnitario": "G",
    "planilhaPrecoTotal": "H",
    "planilhaPrecoUnitarioCopiar": "K",
    "planilhaFator": "RESUMO",
    "BDI": "28.55",
    "colunaFator": "G",
    "linhaFator": "4",
    "planilhaComposicao": "COMPOSICOES",
    "composicaoDescricao": "A",
    "colunaItemDescricaoComposicao": "B",
    "composicaoCoeficiente": "E",
    "composicaoPrecoUnitario": "F",
    "composicaoCoeficienteCopiar": "L",
    "composicaoPrecoUnitarioCopiar": "M",
    "colunaTotaisComposicao": "E",
    "valorTotaisComposicao": "G",
    "planilhaAuxiliar": "COMPOSICOES AUXILIARES",
    "auxiliarDescricao": "A",
    "auxiliarCoeficiente": "E",
    "auxiliarPrecoUnitario": "F",
    "auxiliarCoeficienteCopiar": "L",
    "auxiliarPrecoUnitarioCopiar": "M",
    "colunaTotaisAuxiliar": "E",
    "valorTotaisAuxiliar": "G",
    "valor": "VALOR:",
    "valorComBdi": "VALOR COM BDI"
}


def atualizar_valores_json(
        nome,
        # planilha orcamentaria
        planilhaOrcamentaria,
        colunaInicial,
        colunaFinal,
        valorInicial,
        valorFinal,
        planilhaCodigo,
        planilhaDescricao,
        planilhaQuantidade,
        planilhaPrecoUnitario,
        planilhaPrecoTotal,
        planilhaPrecoUnitarioCopiar,
        # fator
        planilhaFator,
        colunaFator,
        linhaFator,
        bdi,
        # composicao
        planilhaComposicao,
        descricaoComposicao,
        itemDescricaoComposicao,
        coeficienteComposicao,
        precoUnitarioComposicao,
        copiarCoeficienteComposicao,
        copiarPrecoUnitarioComp,
        colunaTotaisComposicao,
        valorTotaisComposicao,
        # aux
        planilhaAuxiliar,
        descricaoAuxiliar,
        coeficienteAuxiliar,
        precoUnitarioAuxiliar,
        copiarCoeficienteAuxiliar,
        copiarPrecoUnitarioAuxiliar,
        colunaTotaisAuxiliar,
        valorTotaisAuxiliar
):
    # Carregar o JSON do arquivo
    with open('config/valores_colunas.json', 'r') as arquivo_json:
        dados_json = json.load(arquivo_json)

    valor_encontrado = False

    # Encontrar o objeto no JSON pelo nome
    for item in dados_json:
        if item["nome"] == nome:
            valor_encontrado = True
            # Atualizar os valores
            # fator
            item["planilhaFator"] = planilhaFator
            item["colunaFator"] = colunaFator
            item["linhaFator"] = linhaFator
            item["BDI"] = bdi
            # planilha orcamentaria
            item["planilhaOrcamentaria"] = planilhaOrcamentaria
            item["colunaInicial"] = colunaInicial
            item["colunaFinal"] = colunaFinal
            item["valorInicial"] = valorInicial
            item["valorFinal"] = valorFinal
            item["planilhaCodigo"] = planilhaCodigo
            item["planilhaDescricao"] = planilhaDescricao
            item["planilhaQuantidade"] = planilhaQuantidade
            item["planilhaPrecoUnitario"] = planilhaPrecoUnitario
            item["planilhaPrecoTotal"] = planilhaPrecoTotal
            item["planilhaPrecoUnitarioCopiar"] = planilhaPrecoUnitarioCopiar
            # composicao
            item["planilhaComposicao"] = planilhaComposicao
            item["composicaoDescricao"] = descricaoComposicao
            item["colunaItemDescricaoComposicao"] = itemDescricaoComposicao
            item["composicaoCoeficiente"] = coeficienteComposicao
            item["composicaoPrecoUnitario"] = precoUnitarioComposicao
            item["composicaoCoeficienteCopiar"] = copiarCoeficienteComposicao
            item["composicaoPrecoUnitarioCopiar"] = copiarPrecoUnitarioComp
            item["colunaTotaisComposicao"] = colunaTotaisComposicao
            item["valorTotaisComposicao"] = valorTotaisComposicao
            # aux
            item["planilhaAuxiliar"] = planilhaAuxiliar
            item["auxiliarDescricao"] = descricaoAuxiliar
            item["auxiliarCoeficiente"] = coeficienteAuxiliar
            item["auxiliarPrecoUnitario"] = precoUnitarioAuxiliar
            item["auxiliarCoeficienteCopiar"] = copiarCoeficienteAuxiliar
            item["auxiliarPrecoUnitarioCopiar"] = copiarPrecoUnitarioAuxiliar
            item["colunaTotaisAuxiliar"] = colunaTotaisAuxiliar
            item["valorTotaisAuxiliar"] = valorTotaisAuxiliar
            break

    if not valor_encontrado:
        item = {
            "nome": nome,
            # fator
            "planilhaFator": planilhaFator,
            "colunaFator": colunaFator,
            "linhaFator": linhaFator,
            "BDI": bdi,
            # planilha orcamentaria
            "planilhaOrcamentaria": planilhaOrcamentaria,
            "colunaInicial": colunaInicial,
            "colunaFinal": colunaFinal,
            "valorInicial": valorInicial,
            "valorFinal": valorFinal,
            "planilhaCodigo": planilhaCodigo,
            "planilhaDescricao": planilhaDescricao,
            "planilhaQuantidade": planilhaQuantidade,
            "planilhaPrecoUnitario": planilhaPrecoUnitario,
            "planilhaPrecoTotal": planilhaPrecoTotal,
            "planilhaPrecoUnitarioCopiar": planilhaPrecoUnitarioCopiar,
            # composicao
            "planilhaComposicao": planilhaComposicao,
            "composicaoDescricao": descricaoComposicao,
            "colunaItemDescricaoComposicao": itemDescricaoComposicao,
            "composicaoCoeficiente": coeficienteComposicao,
            "composicaoPrecoUnitario": precoUnitarioComposicao,
            "composicaoCoeficienteCopiar": copiarCoeficienteComposicao,
            "composicaoPrecoUnitarioCopiar": copiarPrecoUnitarioComp,
            "colunaTotaisComposicao": colunaTotaisComposicao,
            "valorTotaisComposicao": valorTotaisComposicao,
            # aux
            "planilhaAuxiliar": planilhaAuxiliar,
            "auxiliarDescricao": descricaoAuxiliar,
            "auxiliarCoeficiente": coeficienteAuxiliar,
            "auxiliarPrecoUnitario": precoUnitarioAuxiliar,
            "auxiliarCoeficienteCopiar": copiarCoeficienteAuxiliar,
            "auxiliarPrecoUnitarioCopiar": copiarPrecoUnitarioAuxiliar,
            "colunaTotais": colunaTotaisAuxiliar,
            "valorTotais": valorTotaisAuxiliar,
        }
        dados_json.append(item)

    # Escrever os dados de volta no arquivo
    with open('config/valores_colunas.json', 'w') as arquivo_json:
        json.dump(dados_json, arquivo_json, indent=2)

    messagebox.showinfo('Sucesso', 'Item salvo com sucesso!')


def adicionar_item(self, dropdown_valor, menu_dropdown, frame_menu):
    self.todos_dados.append(new_item)
    dropdown_valor.set('Novo')
    salvar_valores_colunas(self, dropdown_valor)
    atualizar_dropdown(menu_dropdown, frame_menu, dropdown_valor)
    dropdown_valor.set('Novo')


def excluir_item(self, dropdown_valor, menu_dropdown, frame_menu):
    # filtar self.todos_dados onde nome seja diferente de dropdown_valor
    self.todos_dados = [
        item for item in self.todos_dados if item["nome"]
        != dropdown_valor.get()]

    with open('config/valores_colunas.json', 'w',
              encoding='utf-8') as arquivo_json:
        json.dump(self.todos_dados, arquivo_json, ensure_ascii=False, indent=2)

    dropdown_valor.set(self.todos_dados[0]["nome"])
    atualizar_dropdown(menu_dropdown, frame_menu, dropdown_valor)


def salvar_valores_colunas(self, dropdown_valor):
    var_nome = dropdown_valor.get()

    # planilha orcamentaria
    var_planilha_orcamentaria = self.entry_planilha_orcamentaria.get()
    var_coluna_inicial = self.entry_coluna_inicial.get()
    var_coluna_final = self.entry_coluna_final.get()
    var_valor_inicial = self.entry_valor_inicial.get()
    var_valor_final = self.entry_valor_final.get()
    var_plan_codigo = self.entry_planilha_codigo.get()
    var_plan_descricao = self.entry_planilha_descricao.get()
    var_plan_quantidade = self.entry_planilha_quantidade.get()
    var_plan_preco_unitario = self.entry_planilha_preco_unitario.get()
    var_plan_preco_total = self.entry_planilha_preco_total.get()
    var_plan_preco_unitario_copiar = self.entry_plan_preco_unit_copiar.get()

    # fator
    var_planilha_fator = self.entry_planilha_fator.get()
    var_coluna_fator = self.entry_coluna_fator.get()
    var_linha_fator = self.entry_linha_fator.get()
    var_bdi = self.entry_bdi.get()

    # composicao
    var_planilha_composicao = self.entry_planilha_comp.get()
    var_descricao_composicao = self.entry_descricao_comp.get()
    var_item_descricao_composicao = self.entry_item_descricao_comp.get()
    var_coefiente_composicao = self.entry_coeficiente_comp.get()
    var_preco_unitario_composicao = self.entry_preco_unitario_comp.get()
    var_coefiente_copiar_composicao = self.entry_coeficiente_copiar_comp.get()
    var_preco_unit_copiar_composicao = self.entry_preco_unit_copiar_comp.get()
    var_coluna_totais_composicao = self.entry_coluna_totais_comp.get()
    var_valor_totais_composicao = self.entry_valor_totais_comp.get()

    # aux
    var_planilha_auxiliar = self.entry_planilha_aux.get()
    var_descricao_aux = self.entry_descricao_aux.get()
    var_coefiente_aux = self.entry_coeficiente_aux.get()
    var_preco_unitario_aux = self.entry_preco_unitario_aux.get()
    var_coefiente_copiar_aux = self.entry_coeficiente_copiar_aux.get()
    var_preco_unit_copiar_aux = self.entry_preco_unit_copiar_aux.get()
    var_coluna_totais_aux = self.entry_coluna_totais_aux.get()
    var_valor_totais_aux = self.entry_valor_totais_aux.get()

    # Atualizar os valores no arquivo JSON
    atualizar_valores_json(
        var_nome,
        var_planilha_orcamentaria,
        var_coluna_inicial,
        var_coluna_final,
        var_valor_inicial,
        var_valor_final,
        var_plan_codigo,
        var_plan_descricao,
        var_plan_quantidade,
        var_plan_preco_unitario,
        var_plan_preco_total,
        var_plan_preco_unitario_copiar,
        var_planilha_fator,
        var_coluna_fator,
        var_linha_fator,
        var_bdi,
        var_planilha_composicao,
        var_descricao_composicao,
        var_item_descricao_composicao,
        var_coefiente_composicao,
        var_preco_unitario_composicao,
        var_coefiente_copiar_composicao,
        var_preco_unit_copiar_composicao,
        var_coluna_totais_composicao,
        var_valor_totais_composicao,
        var_planilha_auxiliar,
        var_descricao_aux,
        var_coefiente_aux,
        var_preco_unitario_aux,
        var_coefiente_copiar_aux,
        var_preco_unit_copiar_aux,
        var_coluna_totais_aux,
        var_valor_totais_aux
    )

    self.todos_dados = open_valores_colunas()


def interface_menu(self, frame_menu, dropdown_valor):
    frame_menu.pack(pady=10)

    menu_dropdown = tk.OptionMenu(
        frame_menu, dropdown_valor,
        *[d["nome"] for d in self.todos_dados])
    # Configurar a largura do OptionMenu (ajuste conforme necessário)
    menu_dropdown.config(font=(None, 18), width=40)

    # Crie um novo menu com a fonte desejada
    novo_menu = tk.Menu(menu_dropdown, tearoff=0, font=(None, 18))

    # Adicione as opções ao novo menu
    for opcao in self.todos_dados:
        novo_menu.add_command(
            label=opcao.get("nome"), command=lambda opcao=opcao:
                dropdown_valor.set(opcao.get("nome")))

    # Configure o novo menu como o menu do OptionMenu
    menu_dropdown["menu"] = novo_menu

    menu_dropdown.grid(row=0, column=0, padx=10)

    botao_editar = tk.Button(
        frame_menu, text="Editar nome",
        font=(None, 18),
        command=lambda: local_abrir_janela_editar(
            self,
            menu_dropdown,
            frame_menu,
            dropdown_valor
        ))
    botao_editar.grid(row=0, column=1, padx=10)

    botao_adicionar = tk.Button(
        frame_menu, text="Adicionar",
        font=(None, 18),
        command=lambda: adicionar_item(
            self, dropdown_valor, menu_dropdown, frame_menu
        )
    )
    botao_adicionar.grid(row=0, column=2, padx=10)

    botao_salvar = tk.Button(
        frame_menu, text="Salvar dados",
        font=(None, 18),
        command=lambda: salvar_valores_colunas(
            self, dropdown_valor
        ))
    botao_salvar.grid(row=0, column=3, padx=10)

    botao_excluir = tk.Button(
        frame_menu, text="Excluir",
        font=(None, 18),
        command=lambda: excluir_item(
            self, dropdown_valor, menu_dropdown, frame_menu
        )
    )
    botao_excluir.grid(row=0, column=4, padx=10)
