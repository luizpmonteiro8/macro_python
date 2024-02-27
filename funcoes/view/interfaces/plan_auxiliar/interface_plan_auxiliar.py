import tkinter as tk

from funcoes.common.custom_input import custom_input
from funcoes.config.open_config import open_valores_label
from funcoes.get.get_linhas_json import (get_coeficiente_aux,
                                         get_coluna_totais_aux,
                                         get_copiar_coeficiente_aux,
                                         get_copiar_preco_unitario_aux,
                                         get_planilha_aux,
                                         get_preco_unitario_aux,
                                         get_valor_totais_aux)
from funcoes.get.get_valores_label import (
    get_label_coeficiente, get_label_composicao_auxiliar_coeficiente_copiar,
    get_label_composicao_auxiliar_coluna_totais,
    get_label_composicao_auxiliar_preco_unitario_copiar,
    get_label_composicao_auxiliar_valor_totais, get_label_planilha_aux,
    get_label_preco_unitario, get_title_planilha_aux)


def interface_plan_auxiliar(self, frame_aux):
    frame_aux.configure(borderwidth=2, relief="solid")
    frame_aux.pack(padx=10, pady=10)

    valores_label = open_valores_label()

    title_frame_aux = tk.Label(
        frame_aux, text=get_title_planilha_aux(valores_label), font=(None, 18))
    title_frame_aux.grid(row=0, column=0, sticky="w", padx=10)

    # Variáveis StringVar
    var_planilha_auxiliar = tk.StringVar(value=get_planilha_aux(self.dados))
    var_coefiente_aux = tk.StringVar(value=get_coeficiente_aux(self.dados))
    var_preco_unitario_aux = tk.StringVar(
        value=get_preco_unitario_aux(self.dados))
    var_coefiente_copiar_aux = tk.StringVar(
        value=get_copiar_coeficiente_aux(self.dados))
    var_preco_unit_copiar_aux = tk.StringVar(
        value=get_copiar_preco_unitario_aux(self.dados))
    var_coluna_totais_aux = tk.StringVar(
        value=get_coluna_totais_aux(self.dados))
    var_valor_totais_aux = tk.StringVar(
        value=get_valor_totais_aux(self.dados))

    # planilha aux
    self.entry_planilha_aux = custom_input(
        frame_aux,
        get_label_planilha_aux(
            valores_label), var_planilha_auxiliar.get(), row=1)

    # coeficiente
    self.entry_coeficiente_aux = custom_input(
        frame_aux, get_label_coeficiente(
            valores_label), var_coefiente_aux.get(), row=2)

    # preco unitario
    self.entry_preco_unitario_aux = custom_input(
        frame_aux, get_label_preco_unitario(
            valores_label), var_preco_unitario_aux.get(), row=3)

    # coeficiente copiar
    self.entry_coeficiente_copiar_aux = custom_input(
        frame_aux, get_label_composicao_auxiliar_coeficiente_copiar(
            valores_label), var_coefiente_copiar_aux.get(), row=4)

    # preco unitario copiar
    self.entry_preco_unit_copiar_aux = custom_input(
        frame_aux, get_label_composicao_auxiliar_preco_unitario_copiar(
            valores_label), var_preco_unit_copiar_aux.get(), row=5)

    # coluna totias
    self.entry_coluna_totais_aux = custom_input(
        frame_aux, get_label_composicao_auxiliar_coluna_totais(
            valores_label), var_coluna_totais_aux.get(), row=6)

    # valor totais
    self.entry_valor_totais_aux = custom_input(
        frame_aux, get_label_composicao_auxiliar_valor_totais(
            valores_label), var_valor_totais_aux.get(), row=7)
