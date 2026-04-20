import time
import tkinter as tk
import traceback
from tkinter import filedialog

import openpyxl

from funcoes.common.buscar_palavras import buscar_palavra
from funcoes.planilha.funcoes.validar_arquivo_excel import validar_arquivo_excel

from funcoes.planilha.funcoes.verificar_formulas_itens import (
    verificar_e_adicionar_formulas,
)
from funcoes.planilha.funcoes.verificar_adicionar_fator import (
    verificar_e_adicionar_fator,
)
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
from funcoes.view.interfaces.menu.interface_menu import recarregar_entries


def selecionar_arquivo_excel(self):
    if not self.dados:
        tk.messagebox.showwarning(
            "Aviso", "Os dados estão vazios.Abra um arquivo JSON primeiro."
        )
        return

    # Abrir janela para selecionar arquivo Excel
    filepath = filedialog.askopenfilename(
        title="Selecione um arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")],
    )

    if filepath:
        try:
            print(">>> Início do processamento do arquivo")
            start_time = time.time()

            # Mostrar mensagem de processamento
            print(">>> Atualizando label para 'Processando...'")
            self.lbl_processando.config(text="Processando...")
            self.update_idletasks()

            # ============================================
            # VALIDAÇÃO DO ARQUIVO ANTES DO PROCESSAMENTO
            # ============================================
            print(">>> Carregando arquivo Excel e validando...")
            print(
                ">>> Este processo demora de acordo com o tamanho do arquivo. Por favor, aguarde..."
            )
            print(">>> Validando estrutura do arquivo Excel...")
            # valido, workbook, dados_validados = validar_arquivo_excel(
            #     filepath, self.dados
            # )

            # if not valido:
            #     tk.messagebox.showerror(
            #         "Erro de Validação", "Validação falhou ou cancelada pelo usuário."
            #     )
            #     print(">>> ERRO: Validação falhou ou cancelada.")
            #     self.lbl_processando.config(text="Validação falhou!")
            #     return  # PARA o processamento

            # self.dados = (
            #     dados_validados[0]
            #     if isinstance(dados_validados, list)
            #     else dados_validados
            # )
            # self.todos_dados = dados_validados
            # recarregar_entries(self)

            workbook = openpyxl.load_workbook(filepath)
            # recarregar_entries(self)

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
            if linhafinal == -1:
                linhafinal = sheet_planilha.max_row

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
            adicionar_fator_comp(workbook, self.dados)

            print(">>> Calculando custo unitário de execução...")
            custo_unitario_execucao(workbook, self.dados)

            print(">>> Gerando resumo de totais...")
            resumo_totais(workbook, self.dados)

            # ============================================
            # ÚLTIMAS VERIFICAÇÕES - Antes de salvar
            # ============================================
            print(">>> Verificando fórmulas dos itens auxiliares...")
            adicionadas = verificar_e_adicionar_formulas(workbook, self.dados)
            print(f">>> {adicionadas} fórmulas adicionadas")

            print(">>> Verificando e adicionando fatores dos itens...")
            fatores_adicionados = verificar_e_adicionar_fator(workbook, self.todos_item)
            print(f">>> {fatores_adicionados} fatores adicionados")

            # ============================================
            # SALVAR ARQUIVO
            # ============================================
            print(">>> Salvando arquivo...")
            print(">>> Confirme na tela principal se foi salvo com sucesso.")
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

        finally:
            pass

