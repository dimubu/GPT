import importlib
import tkinter as tk
from tkinter import ttk, messagebox

# Dicionário de relatórios disponíveis
relatorios = {
    "1": ("Produção de Camas", "relatorios.producao.processador"),
    "2": ("Tecidos por Cores", "relatorios.tecido.processador"),
    "3": ("Média de Produção", "relatorios.media_producao.processador")
}

def executar_relatorio(codigo):
    nome, path = relatorios[codigo]
    try:
        modulo = importlib.import_module(path)
        modulo.main()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao executar '{nome}':\n{e}")

# Interface principal
janela = tk.Tk()
janela.title("Sistema de Relatórios")
janela.geometry("400x250")
janela.resizable(False, False)

# Título
ttk.Label(janela, text="Selecione um relatório para executar:", font=("Segoe UI", 12)).pack(pady=15)

# Lista de relatórios
frame_botoes = ttk.Frame(janela)
frame_botoes.pack(pady=10)

for codigo, (nome, _) in relatorios.items():
    ttk.Button(frame_botoes, text=nome, width=30, command=lambda c=codigo: executar_relatorio(c)).pack(pady=5)

# Botão sair
ttk.Button(janela, text="Sair", width=15, command=janela.destroy).pack(pady=15)

# Inicia o loop da interface
tk.mainloop()
