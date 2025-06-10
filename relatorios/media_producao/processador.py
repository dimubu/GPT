import pandas as pd
import os
import re
from tkinter import Tk, filedialog, messagebox
from datetime import datetime

MEDIDAS_VALIDAS = ["138", "128", "96", "88", "79", "78", "69"]

def calcular_frequencia_diaria(x):
    if x < 1 and x > 0:
        return f"1 a cada {round(1 / x)} dias"
    elif x >= 1:
        return f"{round(x)}"
    return "N/A"

def calcular_frequencia_mensal(x):
    return f"{round(x)}" if x > 0 else "N/A"

def calcular_media(df, dias):
    df = df[["#TAMANHO_DA_CAMA", "#QUANTIDADE_DE_PRODUCAO"]].copy()
    df = df[df["#TAMANHO_DA_CAMA"].notna()]
    df["#TAMANHO_DA_CAMA"] = df["#TAMANHO_DA_CAMA"].astype(str).str.strip()
    df = df[df["#TAMANHO_DA_CAMA"].isin(MEDIDAS_VALIDAS)]
    df["#QUANTIDADE_DE_PRODUCAO"] = pd.to_numeric(df["#QUANTIDADE_DE_PRODUCAO"], errors="coerce")
    df = df.dropna(subset=["#QUANTIDADE_DE_PRODUCAO"])

    agrupado = df.groupby("#TAMANHO_DA_CAMA")["#QUANTIDADE_DE_PRODUCAO"].sum().reset_index()
    agrupado["MEDIA_DIARIA"] = (agrupado["#QUANTIDADE_DE_PRODUCAO"] / dias).round(2)
    agrupado["MEDIA_MENSAL"] = (agrupado["MEDIA_DIARIA"] * 22).round(2)

    diaria = agrupado[["#TAMANHO_DA_CAMA", "MEDIA_DIARIA"]].copy()
    diaria["FREQUENCIA"] = diaria["MEDIA_DIARIA"].apply(calcular_frequencia_diaria)

    mensal = agrupado[["#TAMANHO_DA_CAMA", "MEDIA_MENSAL"]].copy()
    mensal["FREQUENCIA"] = mensal["MEDIA_MENSAL"].apply(calcular_frequencia_mensal)

    return diaria, mensal

def main():
    root = Tk()
    root.withdraw()
    messagebox.showinfo("Selecionar Arquivo", "Selecione o arquivo PRODUCAO com as abas BOX e BAU")

    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo de produção",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if not caminho_arquivo:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
        return

    match = re.search(r"\[(\d+)]", os.path.basename(caminho_arquivo))
    if not match:
        messagebox.showerror("Erro", "Nome do arquivo não contém número de dias entre colchetes [N].")
        return

    dias = int(match.group(1))
    pasta_relatorio = os.path.join(os.path.dirname(caminho_arquivo), f"RELATORIO_[{dias}]_DIAS_ANALISADOS")
    os.makedirs(pasta_relatorio, exist_ok=True)

    try:
        df_box = pd.read_excel(caminho_arquivo, sheet_name="BOX")
        df_bau = pd.read_excel(caminho_arquivo, sheet_name="BAU")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler as abas do arquivo: {e}")
        return

    media_diaria_box, media_mensal_box = calcular_media(df_box.copy(), dias)
    media_diaria_bau, media_mensal_bau = calcular_media(df_bau.copy(), dias)

    data_hoje = datetime.now().strftime("%d-%m-%y")
    caminho_saida_diaria = os.path.join(pasta_relatorio, f"MEDIA_DIARIA_TAMANHO_{data_hoje}.xlsx")
    caminho_saida_mensal = os.path.join(pasta_relatorio, f"MEDIA_MENSAL_TAMANHO_{data_hoje}.xlsx")

    with pd.ExcelWriter(caminho_saida_diaria, engine="openpyxl") as writer:
        media_diaria_box.to_excel(writer, sheet_name="BOX", index=False)
        media_diaria_bau.to_excel(writer, sheet_name="BAU", index=False)

    with pd.ExcelWriter(caminho_saida_mensal, engine="openpyxl") as writer:
        media_mensal_box.to_excel(writer, sheet_name="BOX", index=False)
        media_mensal_bau.to_excel(writer, sheet_name="BAU", index=False)

    messagebox.showinfo("Finalizado", f"Relatórios salvos na pasta:\n{pasta_relatorio}")

if __name__ == "__main__":
    main()
