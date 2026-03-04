import pandas as pd


def tratar_prestador(caminho):

    df = pd.read_csv(
        caminho,
        sep=",",
        encoding="latin1"
    )

    if df.shape[1] == 1:

        df = df.iloc[:,0].str.split(";", expand=True)

    df.to_csv(
        caminho,
        sep=";",
        index=False,
        encoding="latin1"
    )