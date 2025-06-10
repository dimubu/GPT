import importlib
import argparse
import tkinter as tk
from tkinter import ttk, messagebox

# Dicionário de relatórios disponíveis
relatorios = {
    "1": ("Produção de Camas", "relatorios.producao.processador"),
    "2": ("Tecidos por Cores", "relatorios.tecido.processador"),
    "3": ("Média de Produção", "relatorios.media_producao.processador"),
}

def executar_relatorio(codigo):
    nome, path = relatorios[codigo]
    try:
        modulo = importlib.import_module(path)
        modulo.main()
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao executar '{nome}':\n{e}")


class RelatoriosApp(tk.Tk):
    """Janela principal do sistema de relatórios."""

    def __init__(self):
        super().__init__()
        self.title("Sistema de Relatórios")
        self.geometry("400x250")
        self.resizable(False, False)
        ttk.Style().theme_use("clam")
        self._criar_widgets()

    def _criar_widgets(self):
        ttk.Label(self, text="Selecione um relatório para executar:",
                  font=("Segoe UI", 12)).pack(pady=15)

        frame_botoes = ttk.Frame(self)
        frame_botoes.pack(pady=10)
        for codigo, (nome, _) in relatorios.items():
            ttk.Button(
                frame_botoes,
                text=nome,
                width=30,
                command=lambda c=codigo: executar_relatorio(c),
            ).pack(pady=5)

        ttk.Button(self, text="Sair", width=15, command=self.destroy).pack(pady=15)

def main(test_duration: float | None = None) -> None:
    """Executa a aplicação principal.

    Se ``test_duration`` for informado, a janela é fechada após o tempo
    especificado (em segundos). Isso facilita testes automatizados em
    ambientes sem interface gráfica.
    """

    app = RelatoriosApp()
    if test_duration is not None:
        app.after(int(test_duration * 1000), app.destroy)
    app.mainloop()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Sistema de Relatórios")
    parser.add_argument(
        "--test-duration",
        type=float,
        metavar="SEG",
        help="executa em modo de teste por N segundos",
    )
    args = parser.parse_args()
    main(args.test_duration)
