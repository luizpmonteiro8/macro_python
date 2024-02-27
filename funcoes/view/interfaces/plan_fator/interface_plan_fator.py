import tkinter as tk

from funcoes.common.custom_input import custom_input
from funcoes.config.open_config import open_valores_label
from funcoes.get.get_linhas_json import (get_coluna_fator, get_linha_fator,
                                         get_planilha_fator, get_valor_bdi)
from funcoes.get.get_valores_label import (get_label_coluna_fator,
                                           get_label_linha_fator,
                                           get_label_planilha_fator,
                                           get_title_planilha_fator)


def interface_planilha_fator(self, frame_fator):
    frame_fator.configure(borderwidth=2, relief="solid")
    frame_fator.pack(padx=10, pady=10)

    valores_label = open_valores_label()

    title_frame_fator = tk.Label(frame_fator, text=get_title_planilha_fator(
        valores_label
    ), font=(None, 18))
    title_frame_fator.grid(row=0, column=0, sticky="w", padx=10)

    # Variáveis StringVar
    var_planilha_fator = tk.StringVar(value=get_planilha_fator(self.dados))
    var_bdi = tk.StringVar(value=get_valor_bdi(self.dados))
    var_coluna_fator = tk.StringVar(value=get_coluna_fator(self.dados))
    var_linha_fator = tk.StringVar(value=get_linha_fator(self.dados))

    # planilha fator
    self.entry_planilha_fator = custom_input(
        frame_fator,
        get_label_planilha_fator(
            valores_label), var_planilha_fator.get(), row=1)

    # bdi
    self.entry_bdi = custom_input(
        frame_fator, "BDI", var_bdi.get(), row=2,
    )

    # coluna fator
    self.entry_coluna_fator = custom_input(
        frame_fator, get_label_coluna_fator(
            valores_label), var_coluna_fator.get(), row=3)

    # linha fator
    self.entry_linha_fator = custom_input(
        frame_fator, get_label_linha_fator(
            valores_label), var_linha_fator.get(), row=4)
