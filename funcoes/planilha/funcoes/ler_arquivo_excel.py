import time
import tkinter as tk
import traceback
from tkinter import filedialog

import openpyxl

from funcoes.common.buscar_palavras import buscar_palavra
from funcoes.common.custo_unitario import custo_unitario_execucao
from funcoes.get.get_linhas_json import (
    get_coluna_final,
    get_coluna_inicial,
    get_planilha_orcamentaria,
    get_valor_final,
    get_valor_inicial,
)
from funcoes.planilha.funcoes.adicionar_bdi import adicionar_bdi
from funcoes.planilha.funcoes.adicionar_fator import adicionar_fator
from funcoes.planilha.funcoes.adicionar_fator_aux import adicionar_fator_aux
from funcoes.planilha.funcoes.adicionar_fator_comp import adicionar_fator_comp
from funcoes.planilha.funcoes.formula_planilha import (
    copiar_coluna_planilha,
    formula_planilha,
)
from funcoes.planilha.funcoes.resume import resumo_totais
from funcoes.planilha.salvar.salvar_arquivo import salvar_arquivo


def selecionar_arquivo_excel(self):
    if not self.dados:
        tk.messagebox.showwarning(
            "Aviso", "Os dados estão vazios.Abra um arquivo JSON primeiro."
        )
        return

    try:

        # Abrir janela para selecionar arquivo Excel
        filepath = filedialog.askopenfilename(
            title="Selecione um arquivo Excel",
            filetypes=[("Arquivos Excel", "*.xlsx;*.xls")],
        )

        if filepath:
            start_time = time.time()
            # Mostrar mensagem de processamento
            self.lbl_processando.config(text="Processando...")
            self.update_idletasks()
            # Carregar o arquivo Excel
            workbook = openpyxl.load_workbook(filepath)

            # Realizar operações com o workbook conforme necessário
            sheet_name = get_planilha_orcamentaria(self.dados)
            sheet_planilha = workbook[sheet_name]

            # Obter informações de coluna do JSON
            coluna_inicial = get_coluna_inicial(self.dados)
            valor_inicial = get_valor_inicial(self.dados)
            coluna_final = get_coluna_final(self.dados)
            valor_final = get_valor_final(self.dados)

            linhaIni = buscar_palavra(sheet_planilha, coluna_inicial, valor_inicial) + 1
            linhafinal = buscar_palavra(sheet_planilha, coluna_final, valor_final)

            copiar_coluna_planilha(sheet_planilha, self.dados)

            adicionar_fator(workbook, self.dados)

            adicionar_bdi(workbook, self.dados)

            formula_planilha(workbook, linhaIni, linhafinal, self.dados)

            adicionar_fator_aux(workbook, self.dados)

            adicionar_fator_comp(workbook, self.dados, self.item, linhaIni, linhafinal)

            # somatorio_planilha(sheet_planilha)

            custo_unitario_execucao(workbook, self.dados)

            resumo_totais(workbook, self.dados)

            salvar_arquivo(workbook, filepath)

            # sys.exit()

            end_time = time.time()
            total_time = end_time - start_time
            # Limpar mensagem de processamento
            self.lbl_processando.config(
                text="Arquivo gerado com sucesso!"
                + f"Tempo total de execução: {total_time:.2f} segundos"
            )

    except Exception as e:
        tk.messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
        print(f"Ocorreu um erro: {str(e)}")
        traceback.print_exc()
        self.lbl_processando.config(text="Ocorreu um erro!")
