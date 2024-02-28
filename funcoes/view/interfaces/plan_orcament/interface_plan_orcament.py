import tkinter as tk

from funcoes.common.custom_input import custom_input
from funcoes.config.open_config import open_valores_label
from funcoes.get.get_linhas_json import (get_coluna_final, get_coluna_inicial,
                                         get_planilha_codigo,
                                         get_planilha_descricao,
                                         get_planilha_orcamentaria,
                                         get_planilha_preco_total,
                                         get_planilha_preco_unitario,
                                         get_planilha_preco_unitario_copiar,
                                         get_planilha_quantidade,
                                         get_valor_final, get_valor_inicial)
from funcoes.get.get_valores_label import (get_label_codigo,
                                           get_label_coluna_final,
                                           get_label_coluna_inicial,
                                           get_label_descricao,
                                           get_label_planilha_orcamentaria,
                                           get_label_preco_total,
                                           get_label_preco_unitario,
                                           get_label_preco_unitario_copiar,
                                           get_label_quantidade,
                                           get_label_valor_final,
                                           get_label_valor_inicial,
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
    var_plan_codigo = tk.StringVar(
        value=get_planilha_codigo(self.dados))
    var_plan_descricao = tk.StringVar(
        value=get_planilha_descricao(self.dados))
    var_plan_quantidade = tk.StringVar(
        value=get_planilha_quantidade(self.dados))
    var_plan_preco_unitario = tk.StringVar(
        value=get_planilha_preco_unitario(self.dados))
    var_plan_preco_total = tk.StringVar(
        value=get_planilha_preco_total(self.dados))
    var_plan_preco_unitario_copiar = tk.StringVar(
        value=get_planilha_preco_unitario_copiar(self.dados))

    # planilha orcamentaria
    self.entry_planilha_orcamentaria = custom_input(
        frame_orcamentaria,
        get_label_planilha_orcamentaria(
            valores_label), var_planilha_orcamentaria.get(), row=1)

    # coluna inicial
    self.entry_coluna_inicial = custom_input(
        frame_orcamentaria, get_label_coluna_inicial(valores_label),
        var_coluna_inicial.get(), row=2)

    # coluna final
    self.entry_coluna_final = custom_input(
        frame_orcamentaria, get_label_coluna_final(valores_label),
        var_coluna_final.get(), row=3)

    # valor inicial
    self.entry_valor_inicial = custom_input(
        frame_orcamentaria, get_label_valor_inicial(valores_label),
        var_valor_inicial.get(), row=4)

    # valor final
    self.entry_valor_final = custom_input(
        frame_orcamentaria, get_label_valor_final(valores_label),
        var_valor_final.get(), row=5)

    # codigo
    self.entry_planilha_codigo = custom_input(
        frame_orcamentaria, get_label_codigo(valores_label),
        var_plan_codigo.get(), row=6,
    )

    # descricao
    self.entry_planilha_descricao = custom_input(
        frame_orcamentaria, get_label_descricao(valores_label),
        var_plan_descricao.get(), row=7
    )

    # quantidade
    self.entry_planilha_quantidade = custom_input(
        frame_orcamentaria, get_label_quantidade(valores_label),
        var_plan_quantidade.get(), row=8
    )

    # preco unitario
    self.entry_planilha_preco_unitario = custom_input(
        frame_orcamentaria, get_label_preco_unitario(
            valores_label), var_plan_preco_unitario.get(), row=9,
    )

    # preco total
    self.entry_planilha_preco_total = custom_input(
        frame_orcamentaria, get_label_preco_total(
            valores_label), var_plan_preco_total.get(), row=10,
    )

    # preco unitario copiar
    self.entry_plan_preco_unit_copiar = custom_input(
        frame_orcamentaria, get_label_preco_unitario_copiar(
            valores_label), var_plan_preco_unitario_copiar.get(), row=11,
    )
