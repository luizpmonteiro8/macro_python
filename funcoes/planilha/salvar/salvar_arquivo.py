import os
import tkinter as tk
from datetime import datetime


def salvar_arquivo(workbook, filepath):
    # Cria a pasta 'excel-final' se não existir
    output_folder = os.path.join(os.path.dirname(filepath), 'excel-final')
    os.makedirs(output_folder, exist_ok=True)

    # Obtém o nome do arquivo sem extensão
    filename_without_extension = os.path.splitext(
        os.path.basename(filepath))[0]

    data_hora_atual = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

    # Constrói o caminho para o novo arquivo na pasta 'excel-final'
    output_filepath = os.path.join(
        output_folder, f"{filename_without_extension}{data_hora_atual}.xlsx")

    # Salva as alterações no novo arquivo na pasta 'excel-final'
    workbook.save(output_filepath)

    tk.messagebox.showinfo(
        "Concluído",
        "Concluído com sucesso. Arquivo salvo em: " + output_filepath)
