import pandas as pd
from pathlib import Path


def corrigir_csv(caminho):

    try:

        df = pd.read_csv(
            caminho,
            sep=",",
            encoding="latin1",
            dtype=str
        )

        # se veio tudo numa coluna
        if df.shape[1] == 1:

            df = df.iloc[:, 0].str.split(";", expand=True)

        df.to_csv(
            caminho,
            sep=";",
            index=False,
            encoding="latin1"
        )

        print(f"✅ Corrigido: {caminho.name}")

    except Exception as e:

        print(f"❌ Erro em {caminho.name}: {e}")