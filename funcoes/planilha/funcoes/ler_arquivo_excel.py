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
            print(">>> Início do processamento do arquivo")
            start_time = time.time()

            # Mostrar mensagem de processamento
            print(">>> Atualizando label para 'Processando...'")
            self.lbl_processando.config(text="Processando...")
            self.update_idletasks()

            # Carregar o arquivo Excel
            print(">>> Carregando arquivo Excel...")
            workbook = openpyxl.load_workbook(filepath)

            # Ler nome da planilha
            print(">>> Obtendo nome da planilha orçamentária...")
            sheet_name = get_planilha_orcamentaria(self.dados)
            print(f"--- Planilha encontrada: {sheet_name}")
            sheet_planilha = workbook[sheet_name]

            # Obter colunas e valores do JSON
            print(">>> Obtendo colunas e valores iniciais/finais do JSON...")
            coluna_inicial = get_coluna_inicial(self.dados)
            valor_inicial = get_valor_inicial(self.dados)
            coluna_final = get_coluna_final(self.dados)
            valor_final = get_valor_final(self.dados)
            print(f"--- coluna_inicial={coluna_inicial}, valor_inicial={valor_inicial}")
            print(f"--- coluna_final={coluna_final}, valor_final={valor_final}")

            # Buscar linhas
            print(">>> Buscando linha inicial...")
            linhaIni = buscar_palavra(sheet_planilha, coluna_inicial, valor_inicial) + 1
            print(f"--- linhaIni = {linhaIni}")

            print(">>> Buscando linha final...")
            linhafinal = buscar_palavra(sheet_planilha, coluna_final, valor_final)
            print(f"--- linhaFinal = {linhafinal}")

            # Começo das funções de processamento
            print(">>> Copiando colunas da planilha...")
            copiar_coluna_planilha(sheet_planilha, self.dados)

            print(">>> Adicionando Fator...")
            adicionar_fator(workbook, self.dados)

            print(">>> Adicionando BDI...")
            adicionar_bdi(workbook, self.dados)

            print(">>> Inserindo fórmulas na planilha...")
            formula_planilha(workbook, linhaIni, linhafinal, self.dados)

            print(">>> Adicionando Fator Auxiliar...")
            adicionar_fator_aux(workbook, self.dados)

            print(">>> Adicionando Fator de Composição...")
            adicionar_fator_comp(workbook, self.dados, self.item, linhaIni, linhafinal)

            print(">>> Calculando custo unitário de execução...")
            custo_unitario_execucao(workbook, self.dados)

            print(">>> Gerando resumo de totais...")
            resumo_totais(workbook, self.dados)

            print(">>> Salvando arquivo...")
            salvar_arquivo(workbook, filepath)

            print(
                ">>> PROCESSAMENTO FINALIZADO EM {:.2f} segundos".format(
                    time.time() - start_time
                )
            )

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
