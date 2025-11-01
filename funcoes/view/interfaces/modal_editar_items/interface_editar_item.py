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
    try:
        with open("config/valores_item.json", "r", encoding="utf-8") as arquivo_json:
            dados_json = json.load(arquivo_json)

        item_encontrado = False
        for item in dados_json:
            if item["nome"] == old_name:
                item_encontrado = True
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
                break

        if not item_encontrado:
            # Se não encontrou, cria um novo item
            novo_item = {"nome": var_nome.get()}
            for i, key in enumerate(var_nome_values.keys(), 1):
                new_key = f"item{i}"
                novo_item[new_key] = {
                    "nome": var_nome_values[key].get(),
                    "total": var_total_values[key].get(),
                    "adicionarFator": var_add_fator_values[key].get(),
                    "fatorCoeficiente": var_fator_coef_values[key].get(),
                    "buscarAuxiliar": var_buscar_auxiliar_values[key].get(),
                    "iniciaPor": var_inicia_por_values[key].get(),
                    "naoIniciaPor": var_nao_inicia_por_values[key].get(),
                }
            dados_json.append(novo_item)

        with open("config/valores_item.json", "w", encoding="utf-8") as arquivo_json:
            json.dump(dados_json, arquivo_json, indent=2, ensure_ascii=False)

        self.todos_item = open_valores_item()
        dropdown_valor_item.set(var_nome.get())
        nova_janela.destroy()

        # Mostrar mensagem de sucesso
        show_success("Alterações salvas com sucesso!")

    except Exception as e:
        show_error(f"Erro ao salvar: {str(e)}")


def adicionar_item(self, dropdown_valor_item, select_item, nova_janela):
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

    if not item_keys:
        return

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
    nova_janela.attributes("-fullscreen", True)

    # Container principal
    main_frame = tk.Frame(nova_janela)
    main_frame.pack(fill="both", expand=True)

    # Frame para botões principais (fixo no topo)
    frame_botoes_principais = tk.Frame(main_frame)
    frame_botoes_principais.pack(fill="x", pady=10)

    # Canvas e scrollbar
    canvas_editar = tk.Canvas(main_frame)
    canvas_editar.pack(side="left", fill="both", expand=True)

    scrollbar = ttk.Scrollbar(
        main_frame, orient="vertical", command=canvas_editar.yview
    )
    scrollbar.pack(side="right", fill="y")

    canvas_editar.configure(yscrollcommand=scrollbar.set)
    frame = tk.Frame(canvas_editar)
    canvas_editar.create_window((0, 0), window=frame, anchor="nw")

    # Obter dados do item
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

    # Dicionários para armazenar as variáveis
    var_nome_values = {}
    var_total_values = {}
    var_adicionar_fator_values = {}
    var_fator_coeficiente_values = {}
    var_buscar_auxiliar_values = {}
    var_inicia_por_values = {}
    var_nao_inicia_por_values = {}

    # Botões principais
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
        bg="green",
        fg="white",
    )
    btn_salvar.pack(side="left", padx=5)

    btn_adicionar = tk.Button(
        frame_botoes_principais,
        text="Adicionar item",
        command=lambda: adicionar_item(
            self, dropdown_valor_item, select_item, nova_janela
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
    row_counter = 0

    label_nome = tk.Label(frame, text="Nome", font=(None, 16))
    label_nome.grid(row=row_counter, column=0, sticky="w", padx=10, pady=5)

    entry_nome_titulo = tk.Entry(frame, textvariable=var_nome, font=(None, 16))
    entry_nome_titulo.grid(
        row=row_counter, column=1, columnspan=2, sticky="ew", padx=10, pady=5
    )
    row_counter += 1

    ttk.Separator(frame, orient="horizontal").grid(
        row=row_counter, column=0, columnspan=3, sticky="ew", pady=10
    )
    row_counter += 1

    # Lista para armazenar informações dos itens dinâmicos
    dynamic_items = []

    # Processar cada item
    for key, value in select_item.items():
        if isinstance(value, dict):
            # Criar variáveis para este item
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
            item_frame = tk.Frame(frame, relief="groove", bd=2)
            item_frame.grid(
                row=row_counter, column=0, columnspan=3, sticky="ew", padx=10, pady=5
            )
            row_counter += 1

            # Botões de controle do item
            btn_frame = tk.Frame(item_frame)
            btn_frame.grid(row=0, column=0, rowspan=8, sticky="ns", padx=5, pady=5)

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

            # Campos do item
            current_row = 0

            # Nome do item
            label_item = tk.Label(
                item_frame, text=f"Item {len(var_nome_values)}", font=(None, 14)
            )
            label_item.grid(row=current_row, column=1, sticky="w", padx=5, pady=2)

            entry_nome = tk.Entry(
                item_frame, textvariable=var_nome_values[key], font=(None, 14)
            )
            entry_nome.grid(row=current_row, column=2, sticky="ew", padx=5, pady=2)
            current_row += 1

            # Total
            label_total = tk.Label(item_frame, text="Total", font=(None, 14))
            label_total.grid(row=current_row, column=1, sticky="w", padx=5, pady=2)

            entry_total = tk.Entry(
                item_frame, textvariable=var_total_values[key], font=(None, 14)
            )
            entry_total.grid(row=current_row, column=2, sticky="ew", padx=5, pady=2)
            current_row += 1

            # Adicionar Fator
            label_adicionar_fator = tk.Label(
                item_frame, text="Adicionar Fator", font=(None, 14)
            )
            label_adicionar_fator.grid(
                row=current_row, column=1, sticky="w", padx=5, pady=2
            )

            adicionar_fator_dropdown = ttk.Combobox(
                item_frame,
                values=["Sim", "Não"],
                textvariable=var_adicionar_fator_values[key],
                font=(None, 14),
                state="readonly",
            )
            adicionar_fator_dropdown.grid(
                row=current_row, column=2, sticky="ew", padx=5, pady=2
            )
            current_row += 1

            # Fator Coeficiente
            label_fator_coeficiente = tk.Label(
                item_frame, text="Fator Coeficiente", font=(None, 14)
            )
            label_fator_coeficiente.grid(
                row=current_row, column=1, sticky="w", padx=5, pady=2
            )

            fator_coeficiente_dropdown = ttk.Combobox(
                item_frame,
                values=["Sim", "Não"],
                textvariable=var_fator_coeficiente_values[key],
                font=(None, 14),
                state="readonly",
            )
            fator_coeficiente_dropdown.grid(
                row=current_row, column=2, sticky="ew", padx=5, pady=2
            )
            current_row += 1

            # Texto dinâmico
            texto_fator_label = tk.Label(item_frame, font=(None, 12), fg="blue")

            text_position = current_row

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
                            pady=2,
                        )
                    else:
                        label.grid_forget()

                return atualizar_texto

            atualizar_texto_func = criar_funcao_atualizacao(
                key, texto_fator_label, text_position
            )
            atualizar_texto_func()

            var_adicionar_fator_values[key].trace("w", atualizar_texto_func)
            var_fator_coeficiente_values[key].trace("w", atualizar_texto_func)
            current_row += 1

            # Buscar Auxiliar
            label_buscar_auxiliar = tk.Label(
                item_frame, text="Buscar Auxiliar", font=(None, 14)
            )
            label_buscar_auxiliar.grid(
                row=current_row, column=1, sticky="w", padx=5, pady=2
            )

            buscar_auxiliar_dropdown = ttk.Combobox(
                item_frame,
                values=["Sim", "Não"],
                textvariable=var_buscar_auxiliar_values[key],
                font=(None, 14),
                state="readonly",
            )
            buscar_auxiliar_dropdown.grid(
                row=current_row, column=2, sticky="ew", padx=5, pady=2
            )
            current_row += 1

            # Inicia por
            label_inicia_por = tk.Label(item_frame, text="Inicia por", font=(None, 14))
            label_inicia_por.grid(row=current_row, column=1, sticky="w", padx=5, pady=2)

            entry_inicia_por = tk.Entry(
                item_frame, textvariable=var_inicia_por_values[key], font=(None, 14)
            )
            entry_inicia_por.grid(
                row=current_row, column=2, sticky="ew", padx=5, pady=2
            )
            current_row += 1

            # Não inicia por
            label_nao_inicia_por = tk.Label(
                item_frame, text="Não inicia por", font=(None, 14)
            )
            label_nao_inicia_por.grid(
                row=current_row, column=1, sticky="w", padx=5, pady=2
            )

            entry_nao_inicia_por = tk.Entry(
                item_frame, textvariable=var_nao_inicia_por_values[key], font=(None, 14)
            )
            entry_nao_inicia_por.grid(
                row=current_row, column=2, sticky="ew", padx=5, pady=2
            )

            # Configurar pesos das colunas
            item_frame.columnconfigure(2, weight=1)

    # Configurar o grid do frame principal
    frame.columnconfigure(2, weight=1)

    # Atualizar a região de scroll
    frame.update_idletasks()
    canvas_editar.config(scrollregion=canvas_editar.bbox("all"))

    # Bind do mousewheel
    canvas_editar.bind("<MouseWheel>", lambda e: mousewheel(e, canvas_editar, frame))

    # Protocolo de fechamento
    nova_janela.protocol(
        "WM_DELETE_WINDOW", lambda: fechar_janela(canvas_editar, nova_janela)
    )

    nova_janela.mainloop()


def show_error(message):
    """Função auxiliar para mostrar mensagens de erro"""
    error_window = tk.Toplevel()
    error_window.title("Erro")
    error_window.geometry("400x150")
    error_window.transient()
    error_window.grab_set()

    tk.Label(error_window, text=message, wraplength=380, justify="left").pack(
        expand=True, fill="both", padx=10, pady=10
    )
    tk.Button(error_window, text="OK", command=error_window.destroy, width=10).pack(
        pady=10
    )


def show_success(message):
    """Função auxiliar para mostrar mensagens de sucesso"""
    success_window = tk.Toplevel()
    success_window.title("Sucesso")
    success_window.geometry("400x150")
    success_window.transient()
    success_window.grab_set()

    tk.Label(
        success_window, text=message, wraplength=380, justify="left", fg="green"
    ).pack(expand=True, fill="both", padx=10, pady=10)
    tk.Button(success_window, text="OK", command=success_window.destroy, width=10).pack(
        pady=10
    )
