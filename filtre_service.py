import pandas as pd
from pathlib import Path
import datetime


def filter_ocorrencias():
    hoje = datetime.date(2026, 1, 29)

    PATH_CWD = Path.cwd()
    PATH_DOWNLOADS = PATH_CWD / 'src' / 'downloads'
    PATH_DOWNLOADS.mkdir(parents=True, exist_ok=True)

    PATH_OCORRENCIAS = PATH_DOWNLOADS / 'ocorrenciasSF3.xls'
    PATH_HISTORICO = PATH_DOWNLOADS / "historico_processados.csv"

    NOME_DA_ABA_ALVO = 'JANEIRO 2026'
    COLUNA_DATA_REAL = 'Data'
    COLUNA_CONTRATO = 'Contrato'

    #histótico 
    if PATH_HISTORICO.exists():
        df_historico = pd.read_csv(PATH_HISTORICO)
    else:
        df_historico = pd.DataFrame(
            columns=["Contrato", "Data_Ocorrencia", "Data_Envio"]
        )

    print("Linhas no histórico:", len(df_historico))

    #le excel
    df_ocorrencias = pd.read_excel(
        PATH_OCORRENCIAS,
        sheet_name=NOME_DA_ABA_ALVO,
        header=0
    )

    #filtra data
    df_ocorrencias.dropna(subset=[COLUNA_DATA_REAL], inplace=True)
    df_ocorrencias[COLUNA_DATA_REAL] = pd.to_datetime(
        df_ocorrencias[COLUNA_DATA_REAL],
        errors='coerce'
    )
    print("Data filtrada com sucesso.")

    #criando a chave
    df_ocorrencias[COLUNA_CONTRATO] = df_ocorrencias[COLUNA_CONTRATO].astype(str)

    df_ocorrencias["Data_Key"] = (
        df_ocorrencias[COLUNA_DATA_REAL].dt.date.astype(str)
    )

    df_ocorrencias["CHAVE"] = (
        df_ocorrencias[COLUNA_CONTRATO] + "_" + df_ocorrencias["Data_Key"]
    )

    if not df_historico.empty:
        df_historico["CHAVE"] = (
            df_historico["Contrato"].astype(str)
            + "_"
            + df_historico["Data_Ocorrencia"].astype(str)
        )


    #filtro finalizado
    ocorrencias_de_hoje = (
    ocorrencias_de_hoje
    .loc[~ocorrencias_de_hoje["CHAVE"].isin(df_historico.get("CHAVE", []))]
)


    print("Contratos filtrados com sucesso.")

    #gera excel
    path_excel = PATH_DOWNLOADS / "ocorrencias_filtradas.xlsx"
    ocorrencias_de_hoje.to_excel(path_excel, index=False)
    print("Arquivo Excel gerado com sucesso.")

    #formata data p chata da sabrina
    df_html = ocorrencias_de_hoje.copy()
    df_html['Data'] = df_html['Data'].dt.strftime('%d/%m/%Y')
    print("Data formatada para chata da Sabrina.")

        #salva no historico os contratos enviados
    if not ocorrencias_de_hoje.empty:
        df_novos = pd.DataFrame({
            "Contrato": ocorrencias_de_hoje[COLUNA_CONTRATO].astype(str),
            "Data_Ocorrencia": ocorrencias_de_hoje[COLUNA_DATA_REAL].dt.date.astype(str),
            "Data_Envio": hoje.strftime("%Y-%m-%d")
        })

        df_historico = pd.concat([df_historico, df_novos], ignore_index=True)
        df_historico.to_csv(PATH_HISTORICO, index=False)

        print("Histórico atualizado com sucesso.")
    else:
        print("Nenhuma ocorrência nova para salvar no histórico.")


    return df_html, path_excel
