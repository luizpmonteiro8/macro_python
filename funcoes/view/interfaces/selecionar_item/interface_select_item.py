import json
import tkinter as tk

from funcoes.view.interfaces.modal_editar_items.interface_editar_item import (
    interface_editar_item,
)

new_item = {
    "nome": "novo",
    "item1": {
        "nome": "Material",
        "total": "TOTAL Material:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Não",
    },
    "item2": {
        "nome": "Mão de Obra com Encargos Complementares",
        "total": "TOTAL Mão de Obra com Encargos Complementares:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Sim",
        "buscarAuxiliar": "Sim",
    },
    "item3": {
        "nome": "Equipamento Custo Horário",
        "total": "TOTAL Equipamento Custo Horário:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Sim",
        "buscarAuxiliar": "Não",
    },
    "item4": {
        "nome": "Equipamento",
        "total": "TOTAL Equipamento:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Não",
    },
    "item5": {
        "nome": "Serviço",
        "total": "TOTAL Serviço:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Sim",
    },
    "item6": {
        "nome": "Mão de Obra",
        "total": "TOTAL Mão de Obra:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Sim",
        "buscarAuxiliar": "Não",
    },
    "item7": {
        "nome": "Encargos Complementares",
        "total": "TOTAL Encargos Complementares:",
        "adicionarFator": "Não",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Não",
    },
    "item8": {
        "nome": "Especiais",
        "total": "TOTAL Especiais:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Não",
    },
    "item9": {
        "nome": "MATERIAIS",
        "total": "TOTAL MATERIAIS:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Não",
    },
    "item10": {
        "nome": "SERVIÇOS",
        "total": "TOTAL SERVIÇOS:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Sim",
    },
    "item11": {
        "nome": "OUTROS",
        "total": "TOTAL OUTROS:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Não",
    },
    "item12": {
        "nome": "COTAÇÃO / MAO DE OBRA (C/ ENCARGOS)",
        "total": "TOTAL COTAÇÃO / MAO DE OBRA (C/ ENCARGOS):",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Não",
    },
    "item13": {
        "nome": "Não cadastrado",
        "total": "TOTAL Não cadastrado:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Sim",
        "buscarAuxiliar": "Não",
    },
}


def atualizar_interface_select_item(
    self, menu_dropdown, frame_item, dropdown_valor_item
):
    # atualizar menu_dropdown
    menu_dropdown.destroy()

    # for para limpar frame
    for widget in frame_item.winfo_children():
        widget.destroy()

    dropdown_valor_item.set(self.todos_item[0]["nome"])

    interface_select_item(self, frame_item, dropdown_valor_item)


def excluir_item(self, dropdown_valor_item, menu_dropdown, frame_item):
    # filtar self.todos_item onde nome seja diferente de dropdown_valor_item
    self.todos_item = [
        item for item in self.todos_item if item["nome"] != dropdown_valor_item.get()
    ]
    # salvar json
    with open("config/valores_item.json", "w", encoding="utf-8") as arquivo_json:
        json.dump(self.todos_item, arquivo_json, indent=2, ensure_ascii=False)

    atualizar_interface_select_item(
        self, menu_dropdown, frame_item, dropdown_valor_item
    )


def interface_select_item(self, frame_item, dropdown_valor_item):
    frame_item.pack(pady=10)
    menu_dropdown = tk.OptionMenu(
        frame_item, dropdown_valor_item, *[d["nome"] for d in self.todos_item]
    )
    # Configurar a largura do OptionMenu (ajuste conforme necessário)
    menu_dropdown.config(font=(None, 18), width=40)

    # Crie um novo menu com a fonte desejada
    novo_menu = tk.Menu(menu_dropdown, tearoff=0, font=(None, 18))

    # Adicione as opções ao novo menu
    for opcao in self.todos_item:
        novo_menu.add_command(
            label=opcao.get("nome"),
            command=lambda opcao=opcao: dropdown_valor_item.set(opcao.get("nome")),
        )

    # Configure o novo menu como o menu do OptionMenu
    menu_dropdown["menu"] = novo_menu

    menu_dropdown.grid(row=0, column=0, padx=10)

    btn_editar = tk.Button(
        frame_item,
        text="Editar",
        command=lambda: interface_editar_item(self, dropdown_valor_item),
        font=(None, 18),
    )
    btn_editar.grid(row=0, column=1, pady=10, padx=10)

    btn_adicionar = tk.Button(
        frame_item,
        text="Adicionar",
        command=lambda: interface_editar_item(self, dropdown_valor_item, new_item),
        font=(None, 18),
    )
    btn_adicionar.grid(row=0, column=2, pady=10, padx=10)

    btn_excluir = tk.Button(
        frame_item,
        text="Excluir",
        font=(None, 18),
        command=lambda: excluir_item(
            self,
            dropdown_valor_item,
            menu_dropdown,
            frame_item,
        ),
    )
    btn_excluir.grid(row=0, column=3, pady=10, padx=10)
