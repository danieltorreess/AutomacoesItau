import pandas as pd
from pathlib import Path


# 🔥 Cabeçalho FINAL correto (o que você usa manualmente)
COLUNAS_DAP_CORRETAS = [
    "ID","ATENTO_CADASTRO","NOME_CADASTRO","DTHR_CADASTRO","DTHR_INI_CADASTRO",
    "ID_FILA","FILTRO","EVENTO","FECHAMENTO","GERA_TRANSACAO","DTHR_FUP",
    "FILTRO_DIRECIONADO","EVENTO_SOLUCAO","CPF_CLIENTE","NOME_CLIENTE",
    "EMPRESA","CONTRATO","PLATAFORMA","POSSUI_PLACA","PLACA","PRODUTO",
    "PRODUTO_TAB","TITULARIEDADE","TECNOLOGIA_CARTAO","STATUS_CARTAO",
    "SALDO_DISPONIVEL","NOME_EC","CODIGO_EC","VALOR_TRANSACAO",
    "DT_TRANSACAO","MODALIDADE_TRANSACAO","MCC_ESTABELECIMENTO",
    "CODIGO_AUTORIZACAO","ENTRY_MODE_TRANSACAO","NOME_ANALISTA",
    "DT_TRATATIVA_EVENTO","DECISAO_ANALISE","MOTIVO_AJUSTE",
    "REFERENCIA_EVENTO_SOLUCAO","ESTORNO","DATA_DO_ESTORNO",
    "VALORES_ESTORNADOS","BLOQUEIO_POC","POC","PERIODO",
    "INTERLOCUTOR_CADASTRADO","AR_LOTE","ENDERECO_DE_ENTREGA",
    "DATA_ENTREGA","RECEPTOR","ENTREGA","CONTA","BLOQUEIO_DO_CARTAO",
    "SALDO_RESTANTE","OBSERVACOES","PROTOCOLO_FVS","DT_ABERTURA",
    "DT_TRATAMENTO","CARTAO","CPF","NUMERO_DISCADO","PRIMEIRO_CONTATO",
    "BLOQUEIO","ANALISE","MOTIVO","RESPOSTA","DATA_DA_NOTIFICACAO",
    "COMENTARIO","REFERENCIA_NEO","DT_NEO","REFERENCIA_ORIGEM_FVS",
    "REFERENCIA_REANALISE","DT_FVS","TIPO_SOLICITACAO",
    "QUANTIDADE_FICHAS","CNPJ","RAZAO_SOCIAL","CIDADE","UF",
    "ANALISTA","DT_TRATATIVA","BUREAU","DESCRICAO_PRODUTO",
    "QUANTIDADE_CARTOES","QUANTIDADE_RECARGA","VALOR_RECARGA",
    "VALOR_LIMITE","APROVACAO_ALELO","APROVADOR_ALELO",
    "MES_RECEBIMENTO","EMAIL_CONSUMIDOR","EMAIL_CONSULTOR",
    "RECEPCAO","TRATATIVA","DT_ABERTURA_EVENTO","AGING",
    "INTERLOCUTOR","MOTIVO_TELEGRAMA","PARECER",
    "NOTIFICACAO_ENVIADA","EVENTO_TAB","DATA","SEGMENTO",
    "FORMA_PAGAMENTO","BANCO_VENDEDOR","ANALISE_PAGAMENTO",
    "DATA_TRATAMENTO","SITUACAO","CODIGO_PEDIDO","VIP",
    "DATA_PEDIDO","DATA_ULTIMO_EVENTO","INTERLOCUTOR_1",
    "INTERLOCUTOR_2","INTERLOCUTOR_3","AR_MESTRE",
    "STATUS_LOTE","QNTD_CARTOES","EVENTO_1645_REFERENCIA",
    "EVENTO_1522_REFERENCIA","DATA_ABERTURA","DATA_PERIODO",
    "DATA_EXTRACAO","TELEFONE_DISCADO","HISTORICO_ATENDIMENTO",
    "SLA","MOTIVO_EVENTO_INCORRETO",
    "OBSERVACAO_EVENTO_INCORRETO","CARTAO_222",
    "ACESSO_A_COMPRA_E_COMMERCE","PRAZO","DIAS_VENCIMENTO"
]


def processar_dap(caminho_arquivo: Path):

    print(f"\n🔴 Processando DAP: {caminho_arquivo}")

    xls = pd.ExcelFile(caminho_arquivo)
    sheet_original = xls.sheet_names[0]

    df = pd.read_excel(
        caminho_arquivo,
        sheet_name=sheet_original,
        dtype=str
    )

    print(f"📈 Linhas: {len(df)}")
    print(f"📊 Colunas recebidas: {len(df.columns)}")
    print(f"📊 Colunas esperadas: {len(COLUNAS_DAP_CORRETAS)}")

    # 🔥 Validação simples
    if len(df.columns) != len(COLUNAS_DAP_CORRETAS):
        raise Exception(
            f"❌ Estrutura inesperada. "
            f"Recebido: {len(df.columns)} colunas | "
            f"Esperado: {len(COLUNAS_DAP_CORRETAS)}"
        )

    # 🔥 Substitui cabeçalho inteiro
    df.columns = COLUNAS_DAP_CORRETAS

    print("🧾 Cabeçalho ajustado.")

    # 🔥 Limpeza ID_FILA
    if "ID_FILA" in df.columns:
        antes = df["ID_FILA"].astype(str).str.count(r"\?").sum()
        df["ID_FILA"] = df["ID_FILA"].astype(str).str.replace("?", "", regex=False)
        print(f"🧹 ID_FILA limpo. Total de '?' removidos: {antes}")

    # 🔥 Limpeza datas 1900
    for coluna in ["DT_ABERTURA_EVENTO", "DATA_TRATAMENTO"]:
        if coluna in df.columns:
            qtd = df[coluna].astype(str).str.contains("1900", na=False).sum()
            df.loc[df[coluna].astype(str).str.contains("1900", na=False), coluna] = ""
            print(f"🧹 {coluna}: {qtd} datas 1900 removidas.")

    with pd.ExcelWriter(
        caminho_arquivo,
        engine="openpyxl",
        mode="w"
    ) as writer:
        df.to_excel(writer, sheet_name=sheet_original, index=False)

    print("✅ DAP tratado com sucesso.")

    return caminho_arquivo
