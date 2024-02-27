import tkinter as tk

from funcoes.common.custom_input import custom_input
from funcoes.config.open_config import open_valores_label
from funcoes.get.get_linhas_json import (get_coluna_final, get_coluna_inicial,
                                         get_planilha_orcamentaria,
                                         get_valor_final, get_valor_inicial)
from funcoes.get.get_valores_label import (get_label_planilha_orcamentaria,
                                           get_title_planilha_orcamentaria)


def interface_plan_orcamentaria(self, frame_orcamentaria):
    frame_orcamentaria.configure(borderwidth=2, relief="solid")
    frame_orcamentaria.pack(padx=10, pady=10)

    valores_label = open_valores_label()

    title_frame_fator = tk.Label(
        frame_orcamentaria,
        text=get_title_planilha_orcamentaria(
            valores_label
        ), font=(None, 18))
    title_frame_fator.grid(row=0, column=0, sticky="w", padx=10)

    # Variáveis StringVar
    var_planilha_orcamentaria = tk.StringVar(
        value=get_planilha_orcamentaria(self.dados))
    var_coluna_inicial = tk.StringVar(value=get_coluna_inicial(self.dados))
    var_coluna_final = tk.StringVar(value=get_coluna_final(self.dados))
    var_valor_inicial = tk.StringVar(value=get_valor_inicial(self.dados))
    var_valor_final = tk.StringVar(value=get_valor_final(self.dados))

    # planilha orcamentaria
    self.entry_planilha_orcamentaria = custom_input(
        frame_orcamentaria,
        get_label_planilha_orcamentaria(
            valores_label), var_planilha_orcamentaria.get(), row=1)

    # coluna inicial
    self.entry_coluna_inicial = custom_input(
        frame_orcamentaria, "Coluna inicial", var_coluna_inicial.get(), row=2)

    # coluna final
    self.entry_coluna_final = custom_input(
        frame_orcamentaria, "Coluna final", var_coluna_final.get(), row=3)

    # valor inicial
    self.entry_valor_inicial = custom_input(
        frame_orcamentaria, "Valor inicial", var_valor_inicial.get(), row=4)

    # valor final
    self.entry_valor_final = custom_input(
        frame_orcamentaria, "Valor final", var_valor_final.get(), row=5)
