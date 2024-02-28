import tkinter as tk

from funcoes.planilha.funcoes.ler_arquivo_excel import selecionar_arquivo_excel


def interface_select_excel(self, frame_arquivo):
    # Botão para selecionar arquivo Excel
    frame_arquivo.pack(pady=10)
    btn_selecionar_excel = tk.Button(
        frame_arquivo, text="Selecionar Arquivo Excel",
        command=lambda: selecionar_arquivo_excel(self), font=(None, 18))
    btn_selecionar_excel.pack(pady=10)
