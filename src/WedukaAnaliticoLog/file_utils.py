import os
import shutil


def gerar_nome_bkp_incremental(pasta_bkp, nome_base):
    """
    Gera nomes do tipo:
    Analitico_Log_de_acesso_diario_1.xlsx
    Analitico_Log_de_acesso_diario_2.xlsx
    ...
    """
    contador = 1

    while True:
        nome = f"{nome_base}_{contador}.xlsx"
        caminho = os.path.join(pasta_bkp, nome)

        if not os.path.exists(caminho):
            return caminho

        contador += 1


def mover_para_bkp(caminho_atual, pasta_bkp, nome_base):
    """
    Move o arquivo atual para a pasta BKP com nome incremental.
    """
    if not os.path.exists(caminho_atual):
        return

    if not os.path.exists(pasta_bkp):
        os.makedirs(pasta_bkp)

    novo_caminho = gerar_nome_bkp_incremental(pasta_bkp, nome_base)
    shutil.move(caminho_atual, novo_caminho)

    print(f"📦 Backup criado: {novo_caminho}")
