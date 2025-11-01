import json
import tkinter as tk
from tkinter import ttk

from funcoes.common.custom_input import custom_input
from funcoes.config.open_config import open_valores_item


def fechar_janela(canvas_editar, nova_janela):
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
    with open("config/valores_item.json", "r", encoding="utf-8") as arquivo_json:
        dados_json = json.load(arquivo_json)

    for item in dados_json:
        if item["nome"] == old_name:
            # Remove todos os itens existentes
            keys_to_remove = [key for key in item if key.startswith("item")]
            for key in keys_to_remove:
                del item[key]

            # Adiciona os itens na nova ordem
            for i, key in enumerate(var_nome_values.keys(), 1):
                new_key = f"item{i}"
                item[new_key] = {
                    "nome": var_nome_values[key].get(),
                    "total": var_total_values[key].get(),
                    "adicionarFator": var_add_fator_values[key].get(),
                    "fatorCoeficiente": var_fator_coef_values[key].get(),
                    "buscarAuxiliar": var_buscar_auxiliar_values[key].get(),
                    "iniciaPor": var_inicia_por_values[key].get(),
                    "naoIniciaPor": var_nao_inicia_por_values[key].get(),
                }
            item["nome"] = var_nome.get()

    with open("config/valores_item.json", "w", encoding="utf-8") as arquivo_json:
        json.dump(dados_json, arquivo_json, indent=2, ensure_ascii=False)

    self.todos_item = open_valores_item()
    dropdown_valor_item.set(var_nome.get())
    nova_janela.destroy()


def adicionar_item(
    self,
    dropdown_valor_item,
    select_item,
    nova_janela,
    frame_itens,
    var_nome_values,
    var_total_values,
    var_adicionar_fator_values,
    var_fator_coeficiente_values,
    var_buscar_auxiliar_values,
    var_inicia_por_values,
    var_nao_inicia_por_values,
    dynamic_items,
):
    # Encontrar o próximo número disponível
    existing_numbers = [
        int(k.replace("item", "")) for k in select_item.keys() if k.startswith("item")
    ]
    next_number = max(existing_numbers) + 1 if existing_numbers else 1
    novo_item_key = f"item{next_number}"

    novo_item_values = {
        "nome": f"Novo Item {next_number}",
        "total": f"TOTAL Novo Item {next_number}:",
        "adicionarFator": "Sim",
        "fatorCoeficiente": "Não",
        "buscarAuxiliar": "Sim",
        "iniciaPor": "",
        "naoIniciaPor": "",
    }
    select_item[novo_item_key] = novo_item_values

    # Atualizar a interface
    reconstruir_interface(self, dropdown_valor_item, select_item, nova_janela)


def excluir_item(self, dropdown_valor_item, select_item, item_key, nova_janela):
    if len([k for k in select_item.keys() if k.startswith("item")]) <= 1:
        show_error("Não é possível excluir o último item.")
        return

    if item_key in select_item:
        del select_item[item_key]
        reconstruir_interface(self, dropdown_valor_item, select_item, nova_janela)


def mover_item(
    self, dropdown_valor_item, select_item, item_key, direction, nova_janela
):
    item_keys = [k for k in select_item.keys() if k.startswith("item")]

    if direction == "up":
        if item_key != item_keys[0]:
            index = item_keys.index(item_key)
            item_keys[index], item_keys[index - 1] = (
                item_keys[index - 1],
                item_keys[index],
            )
    else:  # down
        if item_key != item_keys[-1]:
            index = item_keys.index(item_key)
            item_keys[index], item_keys[index + 1] = (
                item_keys[index + 1],
                item_keys[index],
            )

    # Reorganizar os itens no dicionário
    new_select_item = {"nome": select_item["nome"]}
    for i, key in enumerate(item_keys, 1):
        new_key = f"item{i}"
        new_select_item[new_key] = select_item[key]

    select_item.clear()
    select_item.update(new_select_item)

    reconstruir_interface(self, dropdown_valor_item, select_item, nova_janela)


def reconstruir_interface(self, dropdown_valor_item, select_item, nova_janela):
    # Fechar a janela atual e abrir uma nova com os dados atualizados
    nova_janela.destroy()
    interface_editar_item(self, dropdown_valor_item, select_item)


def interface_editar_item(self, dropdown_valor_item, updated_item=None):
    nova_janela = tk.Toplevel()
    nova_janela.geometry("900x700")
    nova_janela.title("Editar item")
    # iniciar tela cheia
    nova_janela.attributes("-fullscreen", True)

    canvas_editar = tk.Canvas(nova_janela)
    canvas_editar.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(
        nova_janela, orient="vertical", command=canvas_editar.yview
    )
    scrollbar.pack(side="right", fill="y")

    canvas_editar.configure(yscrollcommand=scrollbar.set)
    frame = tk.Frame(canvas_editar)
    canvas_editar.create_window((0, 0), window=frame, anchor="nw")

    select_item = updated_item or next(
        (
            dado
            for dado in self.todos_item
            if dado.get("nome") == dropdown_valor_item.get()
        ),
        None,
    )

    if not select_item:
        show_error("Item não encontrado.")
        nova_janela.destroy()
        return

    old_name = select_item["nome"]
    var_nome = tk.StringVar(value=select_item["nome"])
    var_nome_values = {}
    var_total_values = {}
    var_adicionar_fator_values = {}
    var_fator_coeficiente_values = {}
    var_buscar_auxiliar_values = {}
    var_inicia_por_values = {}
    var_nao_inicia_por_values = {}

    # Frame para botões principais
    frame_botoes_principais = tk.Frame(frame)
    frame_botoes_principais.grid(row=0, column=0, columnspan=3, pady=10, sticky="ew")

    btn_salvar = tk.Button(
        frame_botoes_principais,
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
        font=(None, 16),
    )
    btn_salvar.pack(side="left", padx=5)

    btn_adicionar = tk.Button(
        frame_botoes_principais,
        text="Adicionar item",
        command=lambda: adicionar_item(
            self,
            dropdown_valor_item,
            select_item,
            nova_janela,
            frame,
            var_nome_values,
            var_total_values,
            var_adicionar_fator_values,
            var_fator_coeficiente_values,
            var_buscar_auxiliar_values,
            var_inicia_por_values,
            var_nao_inicia_por_values,
            [],
        ),
        font=(None, 16),
    )
    btn_adicionar.pack(side="left", padx=5)

    btn_fechar = tk.Button(
        frame_botoes_principais,
        text="Fechar sem Salvar",
        command=lambda: fechar_janela(canvas_editar, nova_janela),
        font=(None, 16),
        fg="white",
        bg="red",
    )
    btn_fechar.pack(side="left", padx=5)

    # Nome do item principal
    entry_nome_titulo = custom_input(frame, "Nome", var_nome.get(), row=1)
    entry_nome_titulo.bind(
        "<KeyRelease>", lambda e, v=var_nome: v.set(entry_nome_titulo.get())
    )

    ttk.Separator(frame, orient="horizontal").grid(
        row=2, column=0, columnspan=3, sticky="ew", pady=10
    )

    row_counter = 3
    items = 1

    # Lista para armazenar informações dos itens dinâmicos
    dynamic_items = []

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

            # Frame para cada item
            item_frame = tk.Frame(frame, relief="groove", bd=1)
            item_frame.grid(
                row=row_counter, column=0, columnspan=3, sticky="ew", padx=5, pady=2
            )

            # Botões de controle do item (mover para cima/baixo, excluir)
            btn_frame = tk.Frame(item_frame)
            btn_frame.grid(row=0, column=0, rowspan=8, sticky="ns", padx=5)

            btn_mover_cima = tk.Button(
                btn_frame,
                text="↑",
                font=(None, 12),
                width=3,
                command=lambda k=key: mover_item(
                    self, dropdown_valor_item, select_item, k, "up", nova_janela
                ),
            )
            btn_mover_cima.pack(pady=2)

            btn_mover_baixo = tk.Button(
                btn_frame,
                text="↓",
                font=(None, 12),
                width=3,
                command=lambda k=key: mover_item(
                    self, dropdown_valor_item, select_item, k, "down", nova_janela
                ),
            )
            btn_mover_baixo.pack(pady=2)

            btn_excluir = tk.Button(
                btn_frame,
                text="X",
                font=(None, 12),
                width=3,
                fg="red",
                command=lambda k=key: excluir_item(
                    self, dropdown_valor_item, select_item, k, nova_janela
                ),
            )
            btn_excluir.pack(pady=2)

            # Calcular posições fixas para cada elemento deste item
            current_rows = {
                "item": 0,
                "total": 1,
                "adicionar_fator": 2,
                "fator_coeficiente": 3,
                "texto_dinamico": 4,
                "buscar_auxiliar": 5,
                "inicia_por": 6,
                "nao_inicia_por": 7,
            }

            entry_nome = custom_input(
                item_frame,
                f"Item {items}",
                var_nome_values[key].get(),
                row=current_rows["item"],
            )
            entry_nome.grid(column=1, sticky="ew", padx=5)
            entry_nome.bind(
                "<KeyRelease>",
                lambda e, v=var_nome_values[key], e2=entry_nome: v.set(e2.get()),
            )

            total_entry = custom_input(
                item_frame,
                "Total",
                var_total_values[key].get(),
                row=current_rows["total"],
            )
            total_entry.grid(column=1, sticky="ew", padx=5)
            total_entry.bind(
                "<KeyRelease>",
                lambda e, v=var_total_values[key], e2=total_entry: v.set(e2.get()),
            )

            # Adicionar Fator
            label_adicionar_fator = tk.Label(
                item_frame, text="Adicionar Fator", font=(None, 16)
            )
            label_adicionar_fator.grid(
                row=current_rows["adicionar_fator"], column=1, sticky="w", padx=5
            )

            adicionar_fator_dropdown = ttk.Combobox(
                item_frame,
                values=["Sim", "Não"],
                textvariable=var_adicionar_fator_values[key],
                font=(None, 16),
            )
            adicionar_fator_dropdown.grid(
                row=current_rows["adicionar_fator"], column=2, sticky="ew", padx=5
            )
            adicionar_fator_dropdown.set(var_adicionar_fator_values[key].get())
            adicionar_fator_dropdown.config(state="readonly")

            # Fator Coeficiente
            label_fator_coeficiente = tk.Label(
                item_frame, text="Fator Coeficiente", font=(None, 16)
            )
            label_fator_coeficiente.grid(
                row=current_rows["fator_coeficiente"], column=1, sticky="w", padx=5
            )

            fator_coeficiente_dropdown = ttk.Combobox(
                item_frame,
                values=["Sim", "Não"],
                textvariable=var_fator_coeficiente_values[key],
                font=(None, 16),
            )
            fator_coeficiente_dropdown.grid(
                row=current_rows["fator_coeficiente"], column=2, sticky="ew", padx=5
            )
            fator_coeficiente_dropdown.set(var_fator_coeficiente_values[key].get())
            fator_coeficiente_dropdown.config(state="readonly")

            # Texto dinâmico - sempre criado mas visível condicionalmente
            texto_fator_label = tk.Label(item_frame, font=(None, 12))

            # Função de atualização com todas as referências necessárias
            def criar_funcao_atualizacao(k, label, row_pos):
                def atualizar_texto(*args):
                    if var_adicionar_fator_values[k].get() == "Sim":
                        texto = (
                            "Vai adicionar fator no coeficiente"
                            if var_fator_coeficiente_values[k].get() == "Sim"
                            else "Vai adicionar fator no preço unitário"
                        )
                        label.config(text=texto)
                        label.grid(
                            row=row_pos,
                            column=1,
                            columnspan=2,
                            sticky="w",
                            padx=5,
                            pady=(0, 5),
                        )
                    else:
                        label.grid_forget()

                return atualizar_texto

            # Criar e armazenar a função de atualização
            atualizar_texto_func = criar_funcao_atualizacao(
                key, texto_fator_label, current_rows["texto_dinamico"]
            )

            # Chamar inicialmente para configurar o estado
            atualizar_texto_func()

            # Registrar as traças
            var_adicionar_fator_values[key].trace("w", atualizar_texto_func)
            var_fator_coeficiente_values[key].trace("w", atualizar_texto_func)

            # Armazenar informações para referência
            dynamic_items.append(
                {
                    "key": key,
                    "label": texto_fator_label,
                    "update_func": atualizar_texto_func,
                    "row": current_rows["texto_dinamico"],
                }
            )

            # Buscar Auxiliar
            label_buscar_auxiliar = tk.Label(
                item_frame, text="Buscar Auxiliar", font=(None, 16)
            )
            label_buscar_auxiliar.grid(
                row=current_rows["buscar_auxiliar"], column=1, sticky="w", padx=5
            )

            buscar_auxiliar_dropdown = ttk.Combobox(
                item_frame,
                values=["Sim", "Não"],
                textvariable=var_buscar_auxiliar_values[key],
                font=(None, 16),
            )
            buscar_auxiliar_dropdown.grid(
                row=current_rows["buscar_auxiliar"], column=2, sticky="ew", padx=5
            )
            buscar_auxiliar_dropdown.set(var_buscar_auxiliar_values[key].get())
            buscar_auxiliar_dropdown.config(state="readonly")

            inicia_por_entry = custom_input(
                item_frame,
                "Inicia por:",
                var_inicia_por_values[key].get(),
                row=current_rows["inicia_por"],
            )
            inicia_por_entry.grid(column=1, sticky="ew", padx=5)
            inicia_por_entry.bind(
                "<KeyRelease>",
                lambda _, v=var_inicia_por_values[key], e=inicia_por_entry: v.set(
                    e.get()
                ),
            )

            nao_inicia_por_entry = custom_input(
                item_frame,
                "Nao inicia por:",
                var_nao_inicia_por_values[key].get(),
                row=current_rows["nao_inicia_por"],
            )
            nao_inicia_por_entry.grid(column=1, sticky="ew", padx=5)
            nao_inicia_por_entry.bind(
                "<KeyRelease>",
                lambda _, v=var_nao_inicia_por_values[
                    key
                ], e=nao_inicia_por_entry: v.set(e.get()),
            )

            # Configurar o grid para expandir
            item_frame.columnconfigure(1, weight=1)
            item_frame.columnconfigure(2, weight=1)

            row_counter += 1
            items += 1

    # Configurar o grid do frame principal
    frame.columnconfigure(1, weight=1)
    frame.columnconfigure(2, weight=1)

    frame.update_idletasks()
    canvas_editar.config(scrollregion=canvas_editar.bbox("all"))
    canvas_editar.bind("<MouseWheel>", lambda e: mousewheel(e, canvas_editar, frame))
    nova_janela.protocol(
        "WM_DELETE_WINDOW", lambda: fechar_janela(canvas_editar, nova_janela)
    )

    nova_janela.mainloop()


def show_error(message):
    """Função auxiliar para mostrar mensagens de erro"""
    error_window = tk.Toplevel()
    error_window.title("Erro")
    error_window.geometry("300x100")

    tk.Label(error_window, text=message, wraplength=280).pack(
        expand=True, fill="both", padx=10, pady=10
    )
    tk.Button(error_window, text="OK", command=error_window.destroy).pack(pady=5)
