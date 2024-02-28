import json
import tkinter as tk
from tkinter import messagebox, simpledialog

from funcoes.config.open_config import open_valores_colunas

# cria janela e editar o nome do json


def atualizar_dropdown(
        menu_dropdown,
        frame_menu,
        dropdown_valor):
    # Remover o menu atual
    menu_dropdown.destroy()
    # Recarregar os dados do dropdown com os valores mais recentes
    dados = open_valores_colunas()
    nomes_colunas = [d["nome"] for d in dados]
    # Criar um novo menu dropdown
    dropdown_valor.set(nomes_colunas[0])
    menu_dropdown = tk.OptionMenu(
        frame_menu, dropdown_valor,
        *nomes_colunas)
    # Configurar a largura do OptionMenu (ajuste conforme necessário)
    menu_dropdown.config(font=(None, 18), width=40)
    # Crie um novo menu com a fonte desejada
    novo_menu = tk.Menu(menu_dropdown, tearoff=0, font=(None, 18))
    # Adicione as opções ao novo menu
    for opcao in dados:
        novo_menu.add_command(
            label=opcao.get("nome"), command=lambda opcao=opcao:
                dropdown_valor.set(opcao.get("nome")))

    # Configure o novo menu como o menu do OptionMenu
    menu_dropdown["menu"] = novo_menu

    menu_dropdown.grid(row=0, column=0, padx=10, sticky="w")


def local_abrir_janela_editar(
    self,
    menu_dropdown,
    frame_menu,
    dropdown_valor
):
    nome_coluna = dropdown_valor.get()
    abrir_janela_editar(nome_coluna)
    self.todos_dados = open_valores_colunas()
    atualizar_dropdown(menu_dropdown, frame_menu, dropdown_valor)
    messagebox.showinfo("Atualizado", "Selecione o item novamente")


def salvar_alteracoes(nome_antigo, novo_nome):
    dados = open_valores_colunas()

    for item in dados:
        if item["nome"] == nome_antigo:
            item["nome"] = novo_nome

    with open("config/valores_colunas.json", "w") as file:
        json.dump(dados, file, indent=2)


def abrir_janela_editar(nome_antigo):
    if nome_antigo:
        novo_nome = simpledialog.askstring(
            "Editar Nome", "Digite o novo nome:", initialvalue=nome_antigo)

        if novo_nome:
            salvar_alteracoes(nome_antigo, novo_nome)
