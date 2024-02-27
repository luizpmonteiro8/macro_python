import json
import tkinter as tk

from funcoes.config.open_config import open_valores_colunas
from funcoes.view.interfaces.menu.funcoes.editar_coluna_nome import \
    local_abrir_janela_editar


def atualizar_valores_json(
        nome,
        planilhaFator,
        colunaFator,
        linhaFator,
        planilhaAuxiliar,
        bdi,
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

    # Encontrar o objeto no JSON pelo nome
    for item in dados_json:
        if item["nome"] == nome:
            # Atualizar os valores
            item["planilhaFator"] = planilhaFator
            item["colunaFator"] = colunaFator
            item["linhaFator"] = linhaFator
            item["planilhaAuxiliar"] = planilhaAuxiliar
            item["BDI"] = bdi
            item["auxiliarCoeficiente"] = coeficienteAuxiliar
            item["auxiliarPrecoUnitario"] = precoUnitarioAuxiliar
            item["auxiliarCoeficienteCopiar"] = copiarCoeficienteAuxiliar
            item["auxiliarPrecoUnitarioCopiar"] = copiarPrecoUnitarioAuxiliar
            item["colunaTotais"] = colunaTotaisAuxiliar
            item["valorTotais"] = valorTotaisAuxiliar
            break

    # Escrever os dados de volta no arquivo
    with open('config/valores_colunas.json', 'w') as arquivo_json:
        json.dump(dados_json, arquivo_json, indent=2)


def salvar_valores_colunas(self, dropdown_valor):
    var_nome = dropdown_valor.get()
    var_planilha_fator = self.entry_planilha_fator.get()
    var_coluna_fator = self.entry_coluna_fator.get()
    var_linha_fator = self.entry_linha_fator.get()
    var_planilha_auxiliar = self.entry_planilha_aux.get()
    var_bdi = self.entry_bdi.get()
    var_coefiente_aux = self.entry_coeficiente_aux.get()
    var_preco_unitario_aux = self.entry_preco_unitario_aux.get()
    var_coefiente_copiar_aux = self.entry_coeficiente_copiar_aux.get()
    var_preco_unit_copiar_aux = self.entry_preco_unit_copiar_aux.get()
    var_coluna_totais_aux = self.entry_coluna_totais_aux.get()
    var_valor_totais_aux = self.entry_valor_totais_aux.get()

    # Atualizar os valores no arquivo JSON
    atualizar_valores_json(
        var_nome,
        var_planilha_fator,
        var_coluna_fator,
        var_linha_fator,
        var_planilha_auxiliar,
        var_bdi,
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

    botao_salvar = tk.Button(
        frame_menu, text="Salvar dados",
        font=(None, 18),
        command=lambda: salvar_valores_colunas(
            self, dropdown_valor
        ))
    botao_salvar.grid(row=0, column=2, padx=10)
