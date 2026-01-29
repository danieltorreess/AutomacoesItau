#type:ignore
## -*- coding: utf-8 -*-
r"""
Created on Ago 30 09:45:01 2022.

@author: AB1416392

Pega os arquivos do SFTP - BI_ALELO\ATENDIMENTO_EC que têm o padrão:
    relatorio-atendimentos-001-2022-05.csv

Salva na pasta de Downloads\Alelo_BKO_Atendimento_EC
Deleta os arquivos que tem nomes iguais com final diferente, para pegar apenas um.
Conserta os arquivos para que ele não tenha colunas a mais do que devia em algumas linhas o que gera error no Visual Studio.
    Com o .replace(u'\ufeff', u'\n')

Joga os arquivos para a rede: \\brsbesrv960\Publico\BASE\ALELO\BKO_ATENDIMENTO_EC

Depois que ele rodar é só rodar a Job JOB_ALELO_BKO_ATENDIMENTO_EC.
Para que ele pegue os arquivos do caminho e atualize a [DB_ALELO].dbo.[TB_ODS_ALELO_BKO_ATENDIMENTO_EC].
A job deleta e insere pelo nome do arquivo, então é bom ficar atento a isso e não ter duas com o mesmo nome.
"""

import paramiko
import os
import re
import warnings
import shutil


def get_file_name(dir_, regex, return_list=False, print_=True):
    """Retorna apenas o nome do arquivo que bate com o regex."""
    os.chdir(dir_)
    arquivos_baixados = os.listdir()
    arquivos_baixados.sort(key=os.path.getmtime, reverse=True)
    r = re.compile(regex)
    arq_list = list(filter(r.match, arquivos_baixados))

    assert len(arq_list) > 0, f'Nenhum arquivo com o regex "{regex} foi encontrado no {dir_}"'

    if return_list:
        sw = f'Tem {len(arq_list)} arquivos com regex "{regex}"'
        if print_:
            print(sw)
        return arq_list

    if len(arq_list) > 1:
        sw = f"\nEm {dir_} tem mais de um arquivo com o regex '{regex}'.\n"\
            "O arquivo com a data de modificação mais recente está sendo utilizado.\n"
        warnings.warn(sw)

    arquivo = arq_list[0]
    return arquivo


def delete_folder(dir_, file_re):
    """Deleta todos on arquivos do DIR com o regex."""
    try:
        arquivos = get_file_name(dir_, file_re, return_list=True)
    except AssertionError:
        print("Nenhum arquivo deletado.")
        return

    for file in arquivos:
        os.remove(f"{dir_}\\{file}")
    print(f"Foram deletados {len(arquivos)} no {dir_}.")


def fix_save(contents, dir_, file):
    """Tira os BOM e conserta as linhas ruins."""
    with open(f"{dir_}\\{file}", 'w', encoding="UTF-8-SIG") as f:
        contents = contents.replace(u'\ufeff\ufeff', u'\n')
        contents = contents.replace(u'\ufeff', u'\n')
        f.write(contents)
    print(f"Arquivo alterado {file}.")


def extract_alelo_bko_ec(localpath, arquivo_re, sftp_dados):
    """Baixa os arquivos do SFTP para o localpath."""
    print("Acessando o SFTP.")
    with paramiko.Transport((sftp_dados['host_name'], 22)) as transport:
        transport.connect(None, sftp_dados['username'], sftp_dados['password'])
        extract_sftp(
            transport, sftp_dados['remotepath'],
            localpath, files_regex=arquivo_re)


def extract_sftp(transport, remotepath, localpath, files_regex):
    """Faz a Extração do SFTP dos arquivos com o Regex."""
    # Connect
    with paramiko.SFTPClient.from_transport(transport) as sftp:
        # Pegar os nomes do arquivos
        sftp.chdir(remotepath)
        files = sftp.listdir()
        ftp_compile = re.compile(files_regex)
        arq_list = list(filter(ftp_compile.match, files))

        assert len(arq_list) != 0, 'Não foi encontrado nenhum arquivo'

        # Salvar em Download
        criar_dir(localpath)
        print(f"\nSalvando arquivos do SFTP em {localpath}.")
        for arq in arq_list:
            sftp.get(arq, f"{localpath}\\{arq}")
            print(f"{arq} salvo.")
        print('')


def criar_dir(path):
    """Cria um Dir vazio, se ele já existir deleta todos os arquivos."""
    if not os.path.exists(path):
        os.mkdir(path)
        print(f'Diretório criado: {path}')
    else:
        shutil.rmtree(f'{path}\\')
        os.mkdir(path)
        print(f'Diretório deletado e criado: {path}')


def transform_alelo_bko_ec(path, arquivo_re):
    """Remove os duplicados, Tira os erros do arquivo e sobrepoe ele."""
    os.chdir(path)
    arquivos = get_file_name(path, arquivo_re, return_list=True, print_=False)

    # Remover duplicados
    print("\nRenomeando e removendo duplicados.")
    for arquivo in arquivos:
        if os.path.exists(arquivo):
            arquivos_iguais = get_file_name(path, arquivo[:34], return_list=True, print_=False)
            arquivos_iguais.sort(reverse=True)
            arq_base = arquivo[:34]
            arq_nao_deletar = [arq for arq in arquivos_iguais if arq != arq_base]
            if arq_nao_deletar == []:
                continue
            arq_nao_deletar = arq_nao_deletar[0]
            if len(arquivos_iguais) > 1:
                print(f'Temos {len(arquivos_iguais)} arquivos com o nome base {arq_base}. '
                      'Utilizando o com o número maior.')
                for arq in arquivos_iguais:
                    if arq != arq_nao_deletar:
                        os.remove(arq)
                        print(f"Arquivo deletado: {arq}")
            os.rename(arq_nao_deletar, arq_base)
            print(f'Nome alterado de {arq_nao_deletar} para {arq_base}')

    # Corrigir Erros
    os.chdir(path)
    arquivos = get_file_name(path, arquivo_re, return_list=True, print_=False)
    print("\nCorrigindo possíveis erros.")
    for arquivo in arquivos:
        with open(f"{path}\\{arquivo}", 'rt', encoding="UTF-8-SIG") as file:
            contents = file.read()
        fix_save(contents, path, arquivo)


def load_alelo_bko_ec(from_path, to_path):
    """Carrega os arquivos do {from_path} para o {to_path}."""
    print(f"\nCarregando os arquivos para {to_path}")
    for file in os.listdir(from_path):
        shutil.copy2(f'{from_path}\\{file}', to_path)
        print(f'{file} salvo.')


def etl_alelo_bko_ec(download_path, localpath, arquivo_re, sftp_dados):
    """Le os arquivos com relatorio-atendimentos, salva na maquina, conserta e salva na rede."""
    #extract_alelo_bko_ec(download_path, arquivo_re, sftp_dados)
    transform_alelo_bko_ec(download_path, arquivo_re)
    load_alelo_bko_ec(download_path, localpath)


if __name__ == "__main__":
    localpath = r"\\brsbesrv960\Publico\BASE\ALELO\BKO_ATENDIMENTO_EC"
    download_path = f"{os.environ['USERPROFILE']}\\Downloads\\Alelo_BKO_Atendimento_EC"
    arquivo_re = r'^relatorio-atendimentos-\d{4}-\d{2}(| \(\d*\))\.csv$'
    sftp_dados = {
        "host_name": "sftp-store.atento.com.br",
        "username": "ab1298644",
        "password": "*****"
    }
    sftp_dados["remotepath"] = f"/home_portal/{sftp_dados['username']}/BI_ALELO/ATENDIMENTO_EC"

    etl_alelo_bko_ec(download_path, localpath, arquivo_re, sftp_dados)
    print("\nRodar a Job JOB_ALELO_BKO_ATENDIMENTO_EC")
    pass
