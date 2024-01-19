import json
import tkinter as tk
from tkinter import filedialog


def abrir_json(self):
    # Abrir janela para selecionar arquivo JSON
    caminho_arquivo = filedialog.askopenfilename(
        filetypes=[("Arquivos JSON", "*.json")])

    if caminho_arquivo:
        # Ler dados do arquivo JSON
        with open(caminho_arquivo, 'r', encoding='utf-8') as arquivo:
            self.dados = json.load(arquivo)

        # Preencher a Treeview com os dados
        atualizar_treeview(self)


def adicionar_item(self):
    # Abrir janela para inserir nova chave e valor
    janela_novo_item = tk.Toplevel(self)
    janela_novo_item.title("Adicionar Novo Item")

    lbl_chave = tk.Label(janela_novo_item, text="Chave:")
    lbl_chave.pack(pady=5)
    entry_chave = tk.Entry(janela_novo_item)
    entry_chave.pack(pady=5)

    lbl_valor = tk.Label(janela_novo_item, text="Valor:")
    lbl_valor.pack(pady=5)
    entry_valor = tk.Entry(janela_novo_item)
    entry_valor.pack(pady=5)

    btn_adicionar_novo_item = tk.Button(janela_novo_item, text="Adicionar",
                                        command=lambda:
                                        adicionar_novo_item(
                                            self,
                                            entry_chave.get(),
                                            entry_valor.get(),
                                            janela_novo_item))
    btn_adicionar_novo_item.pack(pady=10)


def adicionar_novo_item(self, chave, valor, janela):
    if chave and valor:
        # Adicionar novo item ao dicionário
        self.dados[chave] = valor

        # Atualizar Treeview
        atualizar_treeview(self)

        # Fechar janela
        janela.destroy()


def editar_item(self):
    # Obter item selecionado na Treeview
    selecionado = self.treeview.selection()
    if selecionado:
        chave_atual = self.treeview.item(selecionado, 'values')[0]
        valor_atual = self.treeview.item(selecionado, 'values')[1]

        # Abrir janela para editar chave e valor
        janela_editar_item = tk.Toplevel(self)
        janela_editar_item.title("Editar Item")
        janela_editar_item.resizable(False, False)
        janela_editar_item.geometry("+%d+%d" % (
            janela_editar_item.winfo_screenwidth() / 2
            - janela_editar_item.winfo_width(
            ) / 2, janela_editar_item.winfo_screenheight() / 2
            - janela_editar_item.winfo_height() / 2))

        lbl_chave = tk.Label(janela_editar_item, text="Chave:")
        lbl_chave.pack(pady=5)
        entry_chave = tk.Entry(janela_editar_item)
        entry_chave.insert(0, chave_atual)
        entry_chave.pack(pady=5)

        lbl_valor = tk.Label(janela_editar_item, text="Valor:")
        lbl_valor.pack(pady=5)
        entry_valor = tk.Entry(janela_editar_item)
        entry_valor.insert(0, valor_atual)
        entry_valor.pack(pady=5)

        btn_editar_item = tk.Button(janela_editar_item, text="Editar",
                                    command=lambda:
                                    editar_item_selecionado(
                                        self,
                                        entry_chave.get(),
                                        entry_valor.get(),
                                        chave_atual,
                                        janela_editar_item))
        btn_editar_item.pack(pady=10)


def editar_item_selecionado(self, nova_chave, novo_valor, chave_atual,
                            janela):
    if nova_chave and novo_valor:
        # Remover item antigo
        del self.dados[chave_atual]

        # Adicionar item editado
        self.dados[nova_chave] = novo_valor

        # Atualizar Treeview
        atualizar_treeview(self)

        # Fechar janela
        janela.destroy()


def remover_item(self):
    # Obter item selecionado na Treeview
    selecionado = self.treeview.selection()
    if selecionado:
        chave = self.treeview.item(selecionado, 'values')[0]
        # Remover item do dicionário
        del self.dados[chave]
        # Atualizar Treeview
        atualizar_treeview(self)


def salvar_json(self):
    # Mapear a ordem dos itens na Treeview para as chaves no JSON
    ordem_treeview = [self.treeview.item(
        item, 'values')[0] for item in self.treeview.get_children()]

    # Reorganizar os dados de acordo com a ordem da Treeview
    dados_ordenados = {chave: self.dados[chave] for chave in ordem_treeview}

    # Abrir janela para salvar arquivo JSON
    caminho_arquivo = filedialog.asksaveasfilename(
        defaultextension=".json",
        filetypes=[("Arquivos JSON", "*.json")])

    if caminho_arquivo:
        # Salvar dados no arquivo JSON
        with open(caminho_arquivo, 'w') as arquivo:
            # Adicionado o parâmetro default
            json.dump(dados_ordenados, arquivo, indent=2, default=str)

        print("Alterações salvas com sucesso em", caminho_arquivo)


def atualizar_treeview(self):
    # Limpar a Treeview
    for item in self.treeview.get_children():
        self.treeview.delete(item)

    # Preencher a Treeview com os dados
    for chave, valor in self.dados.items():
        self.treeview.insert('', 'end', values=(chave, valor))


def subir_item(self):
    selecionado = self.treeview.selection()
    if selecionado:
        item_id = self.treeview.index(selecionado)
        if item_id > 0:
            self.treeview.move(selecionado, '', item_id - 1)


def descer_item(self):
    selecionado = self.treeview.selection()
    if selecionado:
        item_id = self.treeview.index(selecionado)
        if item_id < len(self.treeview.get_children()) - 1:
            self.treeview.move(selecionado, '', item_id + 1)
