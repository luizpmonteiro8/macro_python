import json
import tkinter as tk
from tkinter import ttk

from funcoes.common.custom_input import custom_input
from funcoes.config.open_config import open_valores_item


def fechar_janela(canvas_editar, nova_janela):
    # Desvincula o evento antes de fechar a janela
    canvas_editar.unbind("<MouseWheel>")
    nova_janela.destroy()


def mousewheel(event, canvas, nova_janela):
    if nova_janela is not None:
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


def salvar_alteracao_json(
    self,
    dropdown_valor_item,
    old_name,
    var_nome,
    var_nome_values,
    var_total_values,
    var_add_fator_values,
    var_fator_coef_values,
    var_buscar_auxiliar_values,
    var_inicia_por_values,
    var_nao_inicia_por_values,
    nova_janela,
):
    # Carregar o JSON do arquivo
    with open("config/valores_item.json", "r", encoding="utf-8") as arquivo_json:
        dados_json = json.load(arquivo_json)

    # Encontrar o objeto no JSON pelo nome
    for item in dados_json:
        if item["nome"] == old_name:
            # Atualizar ou adicionar todos os itens conforme var_nome_values
            for key in var_nome_values:
                # Se o item ainda não existe no JSON, criar um novo dicionário
                if key not in item:
                    item[key] = {}

                # Atualizar os valores, sejam novos ou existentes
                item[key]["nome"] = var_nome_values[key].get()
                item[key]["total"] = var_total_values[key].get()
                item[key]["adicionarFator"] = var_add_fator_values[key].get()
                item[key]["fatorCoeficiente"] = var_fator_coef_values[key].get()
                item[key]["buscarAuxiliar"] = var_buscar_auxiliar_values[key].get()
                item[key]["iniciaPor"] = var_inicia_por_values[key].get()
                item[key]["naoIniciaPor"] = var_nao_inicia_por_values[key].get()

            # Atualizar o nome do item principal
            item["nome"] = var_nome.get()

    # Escrever os dados de volta no arquivo
    with open("config/valores_item.json", "w", encoding="utf-8") as arquivo_json:
        json.dump(dados_json, arquivo_json, indent=2, ensure_ascii=False)

    self.todos_item = open_valores_item()
    dropdown_valor_item.set(var_nome.get())

    nova_janela.destroy()


def adicionar_item(self, dropdown_valor_item, select_item, nova_janela):
    # Gera a chave para o novo item
    novo_item_key = f"item{len(select_item)}"
    novo_item_values = {
        "nome": f"Novo Item {len(select_item)}",
        "total": "TOTAL Novo Item:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Sim",
        "iniciaPor": "",
        "naoIniciaPor": "",
    }

    # Adiciona o novo item ao dicionário select_item
    select_item[novo_item_key] = novo_item_values

    nova_janela.destroy()  # Fechar a janela atual
    interface_editar_item(self, dropdown_valor_item, select_item)


def interface_editar_item(self, dropdown_valor_item, updated_item=None):
    nova_janela = tk.Toplevel()
    nova_janela.geometry("800x600")
    nova_janela.title("Editar item")

    # Canvas com barra de rolagem
    canvas_editar = tk.Canvas(nova_janela)
    canvas_editar.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(
        nova_janela, orient="vertical", command=canvas_editar.yview
    )
    scrollbar.pack(side="right", fill="y")

    canvas_editar.configure(yscrollcommand=scrollbar.set)

    # Frame para conter os widgets
    frame = tk.Frame(canvas_editar)

    # Adicionando o frame ao canvas
    canvas_editar.create_window((0, 0), window=frame, anchor="nw")

    select_item = {}

    if updated_item:
        select_item = updated_item
    else:
        select_item = next(
            (
                dado
                for dado in self.todos_item
                if dado.get("nome") == dropdown_valor_item.get()
            ),
            None,
        )

    # Variáveis StringVar
    old_name = select_item["nome"]
    var_nome = tk.StringVar(value=select_item["nome"])
    var_nome_values = {}
    var_total_values = {}
    var_adicionar_fator_values = {}
    var_fator_coeficiente_values = {}
    var_buscar_auxiliar_values = {}
    var_inicia_por_values = {}
    var_nao_inicia_por_values = {}

    # nome
    entry_nome_titulo = custom_input(frame, "Nome", var_nome.get(), row=1)
    entry_nome_titulo.bind(
        "<KeyRelease>", lambda event, v=var_nome: v.set(entry_nome_titulo.get())
    )

    # separator
    ttk.Separator(frame, orient="horizontal").grid(
        row=2, column=0, columnspan=2, sticky="ew", pady=5
    )

    row_counter = 3
    items = 1
    ultima_linha = 0
    for key, value in select_item.items():
        if isinstance(value, dict):
            var_nome_values[key] = tk.StringVar(value=value["nome"])
            var_total_values[key] = tk.StringVar(value=value["total"])
            var_adicionar_fator_values[key] = tk.StringVar(
                value=value["adicionarFator"]
            )
            var_fator_coeficiente_values[key] = tk.StringVar(
                value=value["fatorCoeficiente"]
            )
            var_buscar_auxiliar_values[key] = tk.StringVar(
                value=value["buscarAuxiliar"]
            )
            var_inicia_por_values[key] = tk.StringVar(value=value.get("iniciaPor", ""))
            var_nao_inicia_por_values[key] = tk.StringVar(
                value=value.get("naoIniciaPor", "")
            )

            # nome
            entry_nome = custom_input(
                frame,
                f"Item {items}",
                var_nome_values[key].get(),
                row=row_counter + 1 + ultima_linha,
            )

            entry_nome.bind(
                "<KeyRelease>",
                lambda event, v=var_nome_values[key], e=entry_nome: v.set(e.get()),
            )

            # total
            total_entry = custom_input(
                frame,
                "Total",
                var_total_values[key].get(),
                row=row_counter + 2 + ultima_linha,
            )

            total_entry.bind(
                "<KeyRelease>",
                lambda event, v=var_total_values[key], e=total_entry: v.set(e.get()),
            )

            # adicionarFator
            label_adicionar_fator = tk.Label(
                frame, text="Adicionar Fator", font=(None, 18)
            )
            label_adicionar_fator.grid(
                row=row_counter + 3 + ultima_linha,
                column=0,
                sticky="ew",
            )
            adicionar_fator_dropdown = ttk.Combobox(
                frame,
                values=["Sim", "Não"],
                textvariable=var_adicionar_fator_values[key],
                font=(None, 18),
            )
            adicionar_fator_dropdown.grid(
                row=row_counter + 3 + ultima_linha, column=1, sticky="ew", padx=10
            )
            adicionar_fator_dropdown.set(
                # Definir o valor inicial
                var_adicionar_fator_values[key].get()
            )
            # Não permitir digitar manualmente
            adicionar_fator_dropdown.config(state="readonly")

            # fatorCoeficiente
            label_fator_coeficiente = tk.Label(
                frame, text="Fator Coeficiente", font=(None, 18)
            )
            label_fator_coeficiente.grid(
                row=row_counter + 4 + ultima_linha,
                column=0,
                sticky="ew",
            )
            fator_coeficiente_dropdown = ttk.Combobox(
                frame,
                values=["Sim", "Não"],
                textvariable=var_fator_coeficiente_values[key],
                font=(None, 18),
            )
            fator_coeficiente_dropdown.grid(
                row=row_counter + 4 + ultima_linha, column=1, sticky="ew", padx=10
            )
            fator_coeficiente_dropdown.set(
                # Definir o valor inicial
                var_fator_coeficiente_values[key].get()
            )
            # Não permitir digitar manualmente
            fator_coeficiente_dropdown.config(state="readonly")

            # buscar Auxiliar
            label_buscar_auxiliar = tk.Label(
                frame, text="Buscar Auxiliar", font=(None, 18)
            )

            inicia_por_entry = custom_input(
                frame,
                "Inicia por:",
                var_inicia_por_values[key].get(),
                row=row_counter + 6 + ultima_linha,
            )
            inicia_por_entry.bind(
                "<KeyRelease>",
                lambda _, v=var_inicia_por_values[key], e=inicia_por_entry: v.set(
                    e.get()
                ),
            )
            nao_inicia_por_entry = custom_input(
                frame,
                "Nao inicia por:",
                var_nao_inicia_por_values[key].get(),
                row=row_counter + 7 + ultima_linha,
            )
            nao_inicia_por_entry.bind(
                "<KeyRelease>",
                lambda _, v=var_nao_inicia_por_values[
                    key
                ], e=nao_inicia_por_entry: v.set(e.get()),
            )
            label_buscar_auxiliar = tk.Label(
                frame, text="Buscar Auxiliar", font=(None, 18)
            )
            label_buscar_auxiliar.grid(
                row=row_counter + 5 + ultima_linha,
                column=0,
                sticky="ew",
            )
            buscar_auxiliar_dropdown = ttk.Combobox(
                frame,
                values=["Sim", "Não"],
                textvariable=var_buscar_auxiliar_values[key],
                font=(None, 18),
            )
            buscar_auxiliar_dropdown.grid(
                row=row_counter + 5 + ultima_linha, column=1, sticky="ew", padx=10
            )
            buscar_auxiliar_dropdown.set(
                # Definir o valor inicial
                var_buscar_auxiliar_values[key].get()
            )
            # Não permitir digitar manualmente
            buscar_auxiliar_dropdown.config(state="readonly")

            # separator
            ttk.Separator(frame, orient="horizontal").grid(
                row=row_counter + 8 + ultima_linha, columnspan=2, sticky="ew", pady=5
            )

            row_counter += 1
            items += 1
            ultima_linha += 8

    # botao salvar
    btn_salvar = tk.Button(
        frame,
        text="Salvar",
        command=lambda: salvar_alteracao_json(
            self,
            dropdown_valor_item,
            old_name,
            var_nome,
            var_nome_values,
            var_total_values,
            var_adicionar_fator_values,
            var_fator_coeficiente_values,
            var_buscar_auxiliar_values,
            var_inicia_por_values,
            var_nao_inicia_por_values,
            nova_janela,
        ),
        font=(None, 18),
    )
    btn_salvar.grid(row=0, column=0, pady=10)

    # butao adicionar
    btn_adicionar = tk.Button(
        frame,
        text="Adicionar item",
        command=lambda: adicionar_item(
            self, dropdown_valor_item, select_item, nova_janela
        ),
        font=(None, 18),
    )
    btn_adicionar.grid(row=0, column=1, pady=10)

    frame.update_idletasks()

    # Configurar o tamanho da área de rolagem
    canvas_editar.config(scrollregion=canvas_editar.bbox("all"))

    # Configurar a função de rolagem do canvas
    canvas_editar.bind(
        "<MouseWheel>", lambda event: mousewheel(event, canvas_editar, frame)
    )

    nova_janela.protocol(
        "WM_DELETE_WINDOW", lambda: fechar_janela(canvas_editar, nova_janela)
    )

    nova_janela.mainloop()
