import os
import pandas as pd

from .config import BKP_PATHS
from .logger import log


def extrair_maior_data(caminho):

    with open(caminho, "r", encoding="latin1") as f:

        linhas = f.readlines()

    ultima = linhas[-1]

    data = ultima.split(";")[0].replace('"', '').strip()

    dt = pd.to_datetime(data, dayfirst=True)

    return dt.strftime("%d-%m-%Y")


def mover_para_bkp(tipo, caminho):

    pasta_bkp = BKP_PATHS[tipo]

    data = extrair_maior_data(caminho)

    nome = os.path.basename(caminho)

    nome_bkp = nome.replace(".csv", f"_{data}.csv")

    destino = os.path.join(pasta_bkp, nome_bkp)

    contador = 1

    while os.path.exists(destino):

        destino = os.path.join(
            pasta_bkp,
            nome.replace(".csv", f"_{data}({contador}).csv")
        )

        contador += 1

    os.rename(caminho, destino)

    log(f"📦 Arquivo movido para BKP: {destino}")