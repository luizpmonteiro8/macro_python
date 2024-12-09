import json
import tkinter as tk
from tkinter import messagebox

def open_valores_colunas():
    try:
        with open("config/valores_colunas.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        show_error("Erro: O arquivo 'config/valores_colunas.json' não foi encontrado.")
    except json.JSONDecodeError:
        show_error("Erro: O arquivo 'config/valores_colunas.json' não pode ser lido. Verifique o formato JSON.")
    except Exception as e:
        show_error(f"Erro inesperado ao abrir 'config/valores_colunas.json': {e}")

def open_valores_label():
    try:
        with open("config/valores_label.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        show_error("Erro: O arquivo 'config/valores_label.json' não foi encontrado.")
    except json.JSONDecodeError:
        show_error("Erro: O arquivo 'config/valores_label.json' não pode ser lido. Verifique o formato JSON.")
    except Exception as e:
        show_error(f"Erro inesperado ao abrir 'config/valores_label.json': {e}")

def open_valores_item():
    try:
        with open("config/valores_item.json", "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        show_error("Erro: O arquivo 'config/valores_item.json' não foi encontrado.")
    except json.JSONDecodeError:
        show_error("Erro: O arquivo 'config/valores_item.json' não pode ser lido. Verifique o formato JSON.")
    except Exception as e:
        show_error(f"Erro inesperado ao abrir 'config/valores_item.json': {e}")

def show_error(message):
    # Cria a janela de erro usando o tkinter
    root = tk.Tk()
    root.withdraw()  # Oculta a janela principal
    messagebox.showerror("Erro", message)  # Exibe a caixa de mensagem de erro
    root.destroy()  # Destroi a janela após mostrar a mensagem
