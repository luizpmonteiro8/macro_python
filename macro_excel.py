import tkinter as tk
import sys

sys.setrecursionlimit(3000)

from funcoes.config.open_config import open_valores_colunas, open_valores_item
from funcoes.view.interfaces.menu.interface_menu import interface_menu
from funcoes.view.interfaces.plan_auxiliar.interface_plan_auxiliar import (
    interface_plan_auxiliar,
)
from funcoes.view.interfaces.plan_composicao.interface_plan_composicao import (
    interface_plan_composicao,
)
from funcoes.view.interfaces.plan_fator.interface_plan_fator import (
    interface_planilha_fator,
)
from funcoes.view.interfaces.plan_orcament.interface_plan_orcament import (
    interface_plan_orcamentaria,
)
from funcoes.view.interfaces.selecionar_excel.interface_select_excel import (
    interface_select_excel,
)
from funcoes.view.interfaces.selecionar_item.interface_select_item import (
    interface_select_item,
)


class MacroExcel(tk.Tk):

    def atualizar_dados(
        self, dropdown_valor, frame_fator, frame_aux, frame_orcamentaria, frame_compo
    ):
        # Atualizar self.dados com base no novo valor do dropdown
        nome_selecionado = dropdown_valor.get()
        self.dados = next(
            (dado for dado in self.todos_dados if dado.get("nome") == nome_selecionado),
            None,
        )

        # Limpar o frame_fator
        for widget in frame_fator.winfo_children():
            widget.destroy()

        # Criar uma nova instância
        interface_planilha_fator(self, frame_fator)

        # Limpar frame_aux
        for widget in frame_aux.winfo_children():
            widget.destroy()

        # Criar uma nova instância
        interface_plan_auxiliar(self, frame_aux)

        # Limpar frame_orcamentaria
        for widget in frame_orcamentaria.winfo_children():
            widget.destroy()

        # Criar uma nova instância
        interface_plan_orcamentaria(self, frame_orcamentaria)

        # Limpar frame_compo
        for widget in frame_compo.winfo_children():
            widget.destroy()

        # Criar uma nova instância
        interface_plan_composicao(self, frame_compo)

    def atualizar_item(self, dropdown_valor_item):
        # Atualizar self.dados com base no novo valor do dropdown
        nome_selecionado = dropdown_valor_item.get()
        self.item = next(
            (dado for dado in self.todos_item if dado.get("nome") == nome_selecionado),
            None,
        )

    def __init__(self):
        super().__init__()

        self.title("Macro excel 3.4")
        self.geometry("1200x800")

        self.todos_dados = open_valores_colunas()
        self.todos_item = open_valores_item()

        # Criar um canvas para conter todas as interfaces
        canvas = tk.Canvas(self)
        canvas.pack(side="left", fill="both", expand=True)

        # Criar um frame que será colocado dentro do canvas
        main_frame = tk.Frame(canvas)

        # Criar uma barra de rolagem mestra
        scrollbar = tk.Scrollbar(canvas)
        scrollbar.pack(side="right", fill="y")

        # Configurar a barra de rolagem para rolar o canvas
        scrollbar.config(command=canvas.yview)

        # frames
        frame_menu = tk.Frame(main_frame)
        frame_item = tk.Frame(main_frame)
        frame_arquivo_excel = tk.Frame(main_frame)
        frame_orcamentaria = tk.Frame(main_frame)
        frame_fator = tk.Frame(main_frame)
        frame_compo = tk.Frame(main_frame)
        frame_aux = tk.Frame(main_frame)

        # linha 1 dropdown, editar e salvar - menu

        dropdown_valor_nome_dados = tk.StringVar()
        dropdown_valor_nome_dados.set(self.todos_dados[0].get("nome"))

        interface_menu(self, frame_menu, dropdown_valor_nome_dados)

        # valor selecionado dados
        self.dados = next(
            (
                dado
                for dado in self.todos_dados
                if dado.get("nome") == dropdown_valor_nome_dados.get()
            ),
            None,
        )

        # valor selecionado item
        dropdown_valor_item = tk.StringVar()
        dropdown_valor_item.set(self.todos_item[0].get("nome"))

        # selecionar item
        interface_select_item(self, frame_item, dropdown_valor_item)

        # valor selecionado item
        self.item = next(
            (
                dado
                for dado in self.todos_item
                if dado.get("nome") == dropdown_valor_item.get()
            ),
            None,
        )

        # selecionar arquivo
        interface_select_excel(self, frame_arquivo_excel)

        self.lbl_processando = tk.Label(frame_arquivo_excel, text="", font=(None, 18))
        self.lbl_processando.pack(pady=0)

        # planilha orcamentaria

        interface_plan_orcamentaria(self, frame_orcamentaria)

        # planilha_fator

        interface_planilha_fator(self, frame_fator)

        # planilha compo

        interface_plan_composicao(self, frame_compo)

        #  planilha_auxiliar

        interface_plan_auxiliar(self, frame_aux)

        # vincular a função de callback ao menu dropdown
        dropdown_valor_nome_dados.trace_add(
            "write",
            lambda *args: self.atualizar_dados(
                dropdown_valor_nome_dados,
                frame_fator,
                frame_aux,
                frame_orcamentaria,
                frame_compo,
            ),
        )
        # atualiza item quando ocorrer troca
        dropdown_valor_item.trace_add(
            "write", lambda *args: self.atualizar_item(dropdown_valor_item)
        )

        # Configurar o canvas para o frame principal (centralizado)
        main_frame.update_idletasks()
        main_frame_width = main_frame.winfo_reqwidth()
        main_frame_height = main_frame.winfo_reqheight()

        canvas.update_idletasks()
        canvas_width = canvas.winfo_reqwidth()
        canvas_height = canvas.winfo_reqheight()

        x_center = (canvas_width - main_frame_width) / 2
        y_center = (canvas_height - main_frame_height) / 2

        canvas.create_window(
            x_center,  # Centralizar horizontalmente
            y_center,  # Centralizar verticalmente
            window=main_frame,
            anchor="center",
        )

        canvas.config(scrollregion=canvas.bbox("all"), yscrollcommand=scrollbar.set)

        # Configurar a função de rolagem do canvas
        canvas.bind_all(
            "<MouseWheel>",
            lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"),
        )


interface = MacroExcel()

interface.mainloop()


# pyinstaller --onefile --hide-console=hide-early macro_excel.py
