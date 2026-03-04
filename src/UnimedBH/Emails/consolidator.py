import pandas as pd
from pathlib import Path


class Consolidator:

    def consolidar(self, pasta_origem, caminho_saida):

        arquivos = sorted(Path(pasta_origem).glob("*.csv"))

        dfs = []

        for arq in arquivos:

            try:
                df = pd.read_csv(
                    arq,
                    sep=";",
                    encoding="latin1",
                    low_memory=False
                )

                dfs.append(df)

                print(f"📄 Lido: {arq.name} ({len(df)} linhas)")

            except Exception as e:

                print(f"❌ Erro lendo {arq.name}: {e}")

        if not dfs:
            print("⚠️ Nenhum arquivo encontrado")
            return

        df_final = pd.concat(dfs, ignore_index=True)

        df_final.to_csv(
            caminho_saida,
            sep=";",          # ⭐ separador correto
            index=False,
            encoding="latin1",
            quoting=1         # ⭐ mantém aspas como no original
        )

        print(f"\n✅ Consolidado criado: {caminho_saida}")
        print(f"📊 Total linhas: {len(df_final)}")