import os
import shutil


def gerar_nome_bkp_incremental(pasta_bkp, nome_base):
    contador = 1

    while True:
        nome = f"{nome_base}_{contador}.xlsx"
        caminho = os.path.join(pasta_bkp, nome)

        if not os.path.exists(caminho):
            return caminho

        contador += 1


def mover_para_bkp(caminho_atual, pasta_bkp, nome_base):
    if not os.path.exists(caminho_atual):
        return

    os.makedirs(pasta_bkp, exist_ok=True)

    novo_caminho = gerar_nome_bkp_incremental(pasta_bkp, nome_base)
    shutil.move(caminho_atual, novo_caminho)

    print(f"📦 Backup criado: {novo_caminho}")
