import sys
import tkinter as tk
import traceback
from tkinter import filedialog, ttk

import openpyxl

from funcoes.adicionar_bdi import adicionar_bdi
from funcoes.adicionar_fator import adicionar_fator
from funcoes.adicionar_fator_aux import adicionar_fator_aux
from funcoes.adicionar_fator_comp import adicionar_fator_comp
from funcoes.comum.buscar_palavras import buscar_palavra
from funcoes.formula_planilha import copiar_coluna_planilha, formula_planilha
from funcoes.interface_json import (abrir_json, adicionar_item, descer_item,
                                    editar_item, remover_item, salvar_json,
                                    subir_item)
from funcoes.salvar_arquivo import salvar_arquivo
from funcoes.somatorio_planilha import somatorio_planilha


class InterfaceJSON(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Macro excel 1.04")
        # self.attributes('-fullscreen', True)
        self.geometry("800x600")

        self.dados = {}  # Dicionário para armazenar os dados

        # Botão para abrir arquivo JSON
        self.btn_abrir_json = tk.Button(
            self, text="Abrir JSON", command=lambda: abrir_json(self))
        self.btn_abrir_json.pack(pady=10)

        # Cria um estilo para o Treeview
        estilo_treeview = ttk.Style()
        estilo_treeview.configure("Treeview.Heading", font=('Arial', 18))
        estilo_treeview.configure("Treeview", font=('Arial', 16))

        # Treeview para exibir/alterar dados JSON
        self.treeview = ttk.Treeview(
            self, columns=('Chave', 'Valor'), show='headings')
        self.treeview.heading('Chave', text='Chave', anchor='w')
        self.treeview.heading('Valor', text='Valor', anchor='w')

        # Configurar as colunas com um percentual da largura da tela
        largura_coluna_chave = int(self.winfo_screenwidth() * 0.2)
        largura_coluna_valor = int(self.winfo_screenwidth() * 0.8)
        self.treeview.column('Chave', width=largura_coluna_chave)
        self.treeview.column('Valor', width=largura_coluna_valor)

        self.treeview.pack(pady=10, padx=10)

        # Frame para os botões de ações
        frame_acoes = tk.Frame(self)
        frame_acoes.pack(pady=10)

        # Botão para adicionar item ao JSON
        self.btn_adicionar = tk.Button(
            frame_acoes, text="Adicionar Item",
            command=lambda: adicionar_item(self))
        self.btn_adicionar.pack(side=tk.LEFT, padx=5)

        # Botão para editar item do JSON
        self.btn_editar = tk.Button(
            frame_acoes, text="Editar Item",
            command=lambda: editar_item(self))
        self.btn_editar.pack(side=tk.LEFT, padx=5)

        # Botão para remover item do JSON
        self.btn_remover = tk.Button(
            frame_acoes, text="Remover Item",
            command=lambda: remover_item(self))
        self.btn_remover.pack(side=tk.LEFT, padx=5)

        # Botão para subir/descer item na Treeview
        self.btn_subir = tk.Button(
            frame_acoes, text="Subir Item",
            command=lambda: subir_item(self))
        self.btn_subir.pack(side=tk.LEFT, padx=5)

        self.btn_descer = tk.Button(
            frame_acoes, text="Descer Item",
            command=lambda: descer_item(self))
        self.btn_descer.pack(side=tk.LEFT, padx=5)

        # Botão para salvar alterações JSON
        self.btn_salvar_json = tk.Button(
            self, text="Salvar JSON",
            command=lambda: salvar_json(self))
        self.btn_salvar_json.pack(pady=10)

        # Botão para selecionar arquivo Excel
        self.btn_selecionar_excel = tk.Button(
            self, text="Selecionar Arquivo Excel",
            command=self.selecionar_arquivo_excel)
        self.btn_selecionar_excel.pack(pady=10)

        self.lbl_processando = tk.Label(self, text="")
        self.lbl_processando.pack(pady=0)

        # Botão fechar
        self.btn_fechar = tk.Button(
            self, text="Fechar",
            command=lambda: sys.exit())
        self.btn_fechar.pack(pady=10)

    def selecionar_arquivo_excel(self):
        if not self.dados:
            tk.messagebox.showwarning(
                "Aviso",
                "Os dados estão vazios.Abra um arquivo JSON primeiro.")
            return

        try:

            # Abrir janela para selecionar arquivo Excel
            filepath = filedialog.askopenfilename(
                title="Selecione um arquivo Excel",
                filetypes=[("Arquivos Excel", "*.xlsx;*.xls")],
            )

            if filepath:
                # Mostrar mensagem de processamento
                self.lbl_processando.config(text="Processando...")
                self.update_idletasks()
                # Carregar o arquivo Excel
                workbook = openpyxl.load_workbook(filepath)

                # Realizar operações com o workbook conforme necessário
                sheet_name = self.dados.get(
                    'planilha', 'PLANILHA ORCAMENTARIA')
                sheet_planilha = workbook[sheet_name]

                # Obter informações de coluna do JSON
                coluna_inicial = self.dados.get(
                    'inicioFim', {}).get('colunaInicial', 'A')
                valor_inicial = self.dados.get(
                    'inicioFim', {}).get('valorInicial', 'ITEM')
                coluna_final = self.dados.get(
                    'inicioFim', {}).get('colunaFinal', 'F')
                valor_final = self.dados.get('inicioFim', {}).get(
                    'valorFinal', 'VALOR BDI TOTAL')

                linhaIni = buscar_palavra(
                    sheet_planilha, coluna_inicial, valor_inicial) + 1
                linhafinal = buscar_palavra(
                    sheet_planilha, coluna_final, valor_final)

                copiar_coluna_planilha(sheet_planilha, self.dados)

                adicionar_fator(workbook, self.dados)

                adicionar_bdi(workbook, self.dados)

                formula_planilha(workbook, linhaIni, linhafinal, self.dados)

                adicionar_fator_aux(workbook, self.dados)

                adicionar_fator_comp(workbook, self.dados,
                                     linhaIni, linhafinal)

                somatorio_planilha(sheet_planilha)

                salvar_arquivo(workbook, filepath)

                # sys.exit()

                # Limpar mensagem de processamento
                self.lbl_processando.config(text="Arquivo gerado com sucesso!")

        except Exception as e:
            tk.messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            print(f"Ocorreu um erro: {str(e)}")
            traceback.print_exc()
            self.lbl_processando.config(text="Ocorreu um erro!")


# Instanciar a interface
interface = InterfaceJSON()
interface.mainloop()

# pyinstaller --onefile --hide-console=hide-early macro_excel.py
