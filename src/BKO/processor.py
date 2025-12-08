# src/BKO/processor.py
import pandas as pd
import os


class BKOProcessor:

    def __init__(self):
        self.destino = r"\\BRSBESRV960\publico\REPORTS\001 CARGAS\001 EXCEL\ITAU\SEGUNDO NIVEL"

    # -------------------------------------------------------

    def processar_pos_venda(self, attachment, filename):
        """
        Salva o arquivo Pós Venda removendo SEMPRE as 2 últimas colunas.
        """
        destino_final = os.path.join(self.destino, filename)
        attachment.SaveAsFile(destino_final)

        # Carrega CSV
        df = pd.read_csv(destino_final, sep=";", encoding="latin1")

        # Remove últimas duas colunas
        if len(df.columns) >= 2:
            df = df.iloc[:, :-2]

        # Salva novamente sobre o mesmo arquivo
        df.to_csv(destino_final, sep=";", index=False, encoding="latin1")

        return destino_final

    # -------------------------------------------------------

    def processar_segundo_nivel(self, attachment, filename):
        """
        Abre e salva novamente — SEM ISSO a Job do SQL quebra.
        """
        destino_final = os.path.join(self.destino, filename)
        attachment.SaveAsFile(destino_final)

        # Reabre e salva (correção do bug do Agent)
        df = pd.read_csv(destino_final, sep=";", encoding="latin1")
        df.to_csv(destino_final, sep=";", index=False, encoding="latin1")

        return destino_final
