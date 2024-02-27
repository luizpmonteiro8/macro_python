import tkinter as tk

from funcoes.config.open_config import open_valores_colunas
from funcoes.view.interfaces.menu.interface_menu import interface_menu
from funcoes.view.interfaces.plan_auxiliar.interface_plan_auxiliar import \
    interface_plan_auxiliar
from funcoes.view.interfaces.plan_composicao.interface_plan_composicao import \
    interface_plan_composicao
from funcoes.view.interfaces.plan_fator.interface_plan_fator import \
    interface_planilha_fator
from funcoes.view.interfaces.plan_orcament.interface_plan_orcament import \
    interface_plan_orcamentaria


class MacroExcel(tk.Tk):

    def atualizar_dados(self, dropdown_valor, frame_fator, frame_aux):
        # Atualizar self.dados com base no novo valor do dropdown
        nome_selecionado = dropdown_valor.get()
        self.dados = next(
            (dado for dado in self.todos_dados if
             dado.get('nome') == nome_selecionado), None)

        # Limpar o frame_fator
        for widget in frame_fator.winfo_children():
            widget.destroy()

        # Criar uma nova instância
        interface_planilha_fator(
            self, frame_fator)

        # Limpar frame_aux
        for widget in frame_aux.winfo_children():
            widget.destroy()

        # Criar uma nova instância
        interface_plan_auxiliar(self, frame_aux)

    def __init__(self):
        super().__init__()

        self.title("Macro excel 1.05")
        self.geometry("1200x800")

        self.todos_dados = open_valores_colunas()

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
        frame_orcamentaria = tk.Frame(main_frame)
        frame_fator = tk.Frame(main_frame)
        frame_compo = tk.Frame(main_frame)
        frame_aux = tk.Frame(main_frame)

        # linha 1 dropdown, editar e salvar - menu

        dropdown_valor = tk.StringVar()
        dropdown_valor.set(self.todos_dados[0].get("nome"))

        interface_menu(self, frame_menu, dropdown_valor)

        # valor selecionado
        self.dados = next((dado for dado in self.todos_dados if dado.get(
            'nome') == dropdown_valor.get()), None)

        # planilha orcamentaria

        interface_plan_orcamentaria(self, frame_orcamentaria)

        # planilha_fator

        interface_planilha_fator(
            self, frame_fator)

        # planilha compo

        interface_plan_composicao(self, frame_compo)

        #  planilha_auxiliar

        interface_plan_auxiliar(self, frame_aux)

        # vincular a função de callback ao menu dropdown
        dropdown_valor.trace_add(
            "write", lambda *args:
                self.atualizar_dados(dropdown_valor, frame_fator, frame_aux
                                     ))

        # Configurar o canvas para o frame principal (centralizado)
        main_frame.update_idletasks()
        main_frame_width = main_frame.winfo_reqwidth()
        main_frame_height = main_frame.winfo_reqheight()

        canvas.update_idletasks()
        canvas_width = canvas.winfo_reqwidth()
        canvas_height = canvas.winfo_reqheight()

        x_center = (canvas_width - main_frame_width) / 2
        y_center = (canvas_height - main_frame_height) / 2

        print(canvas_width, canvas_height)
        print(main_frame_width, main_frame_height)
        print(x_center, y_center)

        canvas.create_window(
            x_center,  # Centralizar horizontalmente
            y_center,  # Centralizar verticalmente
            window=main_frame,
            anchor="center"
        )

        canvas.config(
            scrollregion=canvas.bbox("all"), yscrollcommand=scrollbar.set
        )


interface = MacroExcel()
interface.mainloop()
