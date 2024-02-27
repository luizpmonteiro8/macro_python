# input.py
import tkinter as tk


def custom_input(frame, label, initial_value, row):
    fonte = 18

    # Rótulo (Label)
    label_nome = tk.Label(frame, text=label, font=(None, fonte))
    label_nome.grid(row=row, column=0, sticky="w", padx=10)

    # Entrada (Entry)
    valor_planilha_fator = tk.StringVar(value=initial_value)
    entry_nome = tk.Entry(
        frame, textvariable=valor_planilha_fator, font=(None, fonte))
    entry_nome.grid(row=row, column=1, sticky="ew", padx=10)

    return entry_nome
