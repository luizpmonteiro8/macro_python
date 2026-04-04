"""
Componente reutilizável para janela de correção de valores.
TRATAMENTO CORRETO DO BOTÃO CANCELAR.
"""

import tkinter as tk
from tkinter import messagebox


def janela_corrigir_valor(titulo, mensagem, instrucao, valor_atual, valor_default=None):
    """
    Mostra janela para usuário corrigir um valor.
    TRATAMENTO CORRETO: Se usuário clicar Cancelar, retorna (False, None)
    NÃO continua para o próximo passo.

    Args:
        titulo: Título da janela
        mensagem: O que está errado (leigo)
        instrucao: Passo a passo para encontrar o valor no Excel
        valor_atual: Valor atual no JSON (para referência)
        valor_default: Valor sugerido como padrão

    Returns:
        tuple: (True, novo_valor) se confirmado, (False, None) se cancelado
    """
    resultado = {"confirmado": False, "valor": None}

    def on_confirmar():
        resultado["confirmado"] = True
        resultado["valor"] = entry.get()
        janela.destroy()

    def on_cancelar():
        # BOTÃO CANCELAR FUNCIONA CORRETAMENTE
        # Define confirmado como False e retorna None
        resultado["confirmado"] = False
        resultado["valor"] = None
        janela.destroy()

    janela = tk.Toplevel()
    janela.title(titulo)
    janela.geometry("600x400")
    janela.resizable(False, False)
    janela.grab_set()
    janela.focus_force()
    janela.transient()

    mainframe = tk.Frame(janela, padx=20, pady=20)
    mainframe.pack(fill=tk.BOTH, expand=True)

    tk.Label(
        mainframe,
        text=mensagem,
        font=("Arial", 12, "bold"),
        wraplength=550,
        justify=tk.LEFT,
    ).pack(anchor=tk.W)

    tk.Label(mainframe, text="", font=("Arial", 8)).pack()

    tk.Label(mainframe, text="INSTRUÇÕES:", font=("Arial", 10, "bold")).pack(
        anchor=tk.W
    )
    tk.Label(
        mainframe, text=instrucao, font=("Arial", 10), justify=tk.LEFT, wraplength=550
    ).pack(anchor=tk.W)

    tk.Label(mainframe, text="", font=("Arial", 8)).pack()

    tk.Label(
        mainframe,
        text=f"Valor atual no sistema: '{valor_atual}'",
        font=("Arial", 9),
        fg="gray",
    ).pack(anchor=tk.W)

    tk.Label(mainframe, text="Digite o novo valor:", font=("Arial", 10, "bold")).pack(
        anchor=tk.W, pady=(10, 5)
    )

    entry = tk.Entry(mainframe, width=50, font=("Arial", 11))
    if valor_default:
        entry.insert(0, valor_default)
    elif valor_atual:
        entry.insert(0, valor_atual)
    entry.pack(fill=tk.X, pady=(0, 10))
    entry.focus()

    frame_botoes = tk.Frame(mainframe)
    frame_botoes.pack(fill=tk.X)

    btn_cancelar = tk.Button(
        frame_botoes,
        text="Cancelar",
        font=("Arial", 10),
        width=15,
        command=on_cancelar,
    )
    btn_cancelar.pack(side=tk.RIGHT, padx=(10, 0))

    btn_confirmar = tk.Button(
        frame_botoes,
        text="Corrigir",
        font=("Arial", 10, "bold"),
        width=15,
        bg="#4CAF50",
        fg="white",
        command=on_confirmar,
    )
    btn_confirmar.pack(side=tk.RIGHT)

    janela.bind("<Return>", lambda e: on_confirmar())
    janela.bind("<Escape>", lambda e: on_cancelar())

    janela.wait_window()

    return resultado["confirmado"], resultado["valor"]
