import pandas as pd
import os
import re
from tkinter import Tk, filedialog, messagebox
from datetime import datetime

def calcular_frequencia_textual(media):
    if media < 1:
        return f"1 a cada {round(1 / media):,} dias"
    else:
        return str(round(media))

def calcular_media(df, dias):
    df = df[["#TIPO_DE_TECIDO", "#COR_DO_TECIDO", "#QUANTIDADE_DE_PRODUCAO"]].copy()
    df = df.dropna(subset=["#TIPO_DE_TECIDO", "#COR_DO_TECIDO", "#QUANTIDADE_DE_PRODUCAO"])
    df["#QUANTIDADE_DE_PRODUCAO"] = pd.to_numeric(df["#QUANTIDADE_DE_PRODUCAO"], errors="coerce")
    df = df.dropna(subset=["#QUANTIDADE_DE_PRODUCAO"])

    agrupado = df.groupby(["#TIPO_DE_TECIDO", "#COR_DO_TECIDO"])["#QUANTIDADE_DE_PRODUCAO"].sum().reset_index()
    agrupado["MEDIA_DIARIA"] = (agrupado["#QUANTIDADE_DE_PRODUCAO"] / dias).round(2)
    agrupado["MEDIA_MENSAL"] = (agrupado["MEDIA_DIARIA"] * 22).round(2)

    diaria = agrupado[["#TIPO_DE_TECIDO", "#COR_DO_TECIDO", "MEDIA_DIARIA"]].copy()
    mensal = agrupado[["#TIPO_DE_TECIDO", "#COR_DO_TECIDO", "MEDIA_MENSAL"]].copy()

    diaria["FREQUENCIA"] = diaria["MEDIA_DIARIA"].apply(calcular_frequencia_textual)
    mensal["FREQUENCIA"] = mensal["MEDIA_MENSAL"].apply(calcular_frequencia_textual)

    return diaria, mensal

def main():
    root = Tk()
    root.withdraw()
    messagebox.showinfo("Selecionar Arquivo", "Selecione o arquivo PRODUCAO com as abas BOX e BAU")

    caminho_arquivo = filedialog.askopenfilename(
        title="Selecione o arquivo",
        filetypes=[("Planilhas Excel", "*.xlsx")]
    )

    if not caminho_arquivo:
        messagebox.showerror("Erro", "Nenhum arquivo selecionado.")
        return

    match = re.search(r"\[(\d+)]", os.path.basename(caminho_arquivo))
    if not match:
        messagebox.showerror("Erro", "Nome do arquivo não contém [N] com quantidade de dias.")
        return

    dias = int(match.group(1))
    pasta_saida = os.path.join(os.path.dirname(caminho_arquivo), f"RELATORIO_[{dias}]_DIAS_ANALISADOS")
    os.makedirs(pasta_saida, exist_ok=True)

    try:
        df_box = pd.read_excel(caminho_arquivo, sheet_name="BOX")
        df_bau = pd.read_excel(caminho_arquivo, sheet_name="BAU")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler planilha: {e}")
        return

    media_diaria_box, media_mensal_box = calcular_media(df_box, dias)
    media_diaria_bau, media_mensal_bau = calcular_media(df_bau, dias)

    data_hoje = datetime.now().strftime("%d-%m-%y")
    caminho_saida_diaria = os.path.join(pasta_saida, f"MEDIA_DIARIA_TECIDO_{data_hoje}.xlsx")
    caminho_saida_mensal = os.path.join(pasta_saida, f"MEDIA_MENSAL_TECIDO_{data_hoje}.xlsx")
    caminho_saida_geral = os.path.join(pasta_saida, f"MEDIA_GERAL_TECIDO_{data_hoje}.xlsx")

    with pd.ExcelWriter(caminho_saida_diaria, engine="openpyxl") as writer:
        media_diaria_box.to_excel(writer, sheet_name="BOX", index=False)
        media_diaria_bau.to_excel(writer, sheet_name="BAU", index=False)

    with pd.ExcelWriter(caminho_saida_mensal, engine="openpyxl") as writer:
        media_mensal_box.to_excel(writer, sheet_name="BOX", index=False)
        media_mensal_bau.to_excel(writer, sheet_name="BAU", index=False)

    # Concatenar BOX e BAU para gerar MEDIA_GERAL
    df_geral = pd.concat([df_box, df_bau], ignore_index=True)
    media_diaria_geral, media_mensal_geral = calcular_media(df_geral, dias)

    with pd.ExcelWriter(caminho_saida_geral, engine="openpyxl") as writer:
        media_diaria_geral.to_excel(writer, sheet_name="DIARIA", index=False)
        media_mensal_geral.to_excel(writer, sheet_name="MENSAL", index=False)

    messagebox.showinfo("Sucesso", f"Relatórios gerados em:\n{pasta_saida}")

if __name__ == "__main__":
    main()
