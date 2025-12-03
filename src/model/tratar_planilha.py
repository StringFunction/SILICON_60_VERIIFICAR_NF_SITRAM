import pandas as pd
import logging



def tratar_planilha():
    print("Filtrando planilha")
    logging.info("Filtrando planilha")

    # df = pd.read_excel(r"D:\CODIGOS\SILICON_60_VERIIFICAR_NF_SITRAM\notas_para_consultar_sitram.xlsx", dtype=str)
    df = pd.read_excel(r"notas_para_consultar_sitram.xlsx", dtype=str)
    index = len(df.index)
    df = df[:index-1]
    nr_unicos = df["Nr. Documento"].unique()
    df_nr_unicos = df.loc[df["Nr. Documento"] == nr_unicos]
    df_filtrado = df_nr_unicos.loc[df_nr_unicos["UF do Fornecedor"] != "CE"]
    df_filtrado["Status"] = ""

    # df_filtrado.to_excel(r"D:\CODIGOS\SILICON_60_VERIIFICAR_NF_SITRAM\planilha_para_consultar_sitram.xlsx", index=False)
    df_filtrado.to_excel(r"planilha_para_consultar_sitram.xlsx", index=False)
    print("Processo concluido com sucesso")
    logging.info("Processo concluido com sucesso")

if __name__ == "__main__":
    tratar_planilha()
  

    