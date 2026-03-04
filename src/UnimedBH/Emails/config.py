from datetime import datetime, timedelta

# ===== EMAIL =====

OUTLOOK_FOLDER = ["Caixa de Entrada", "GENESYS"]

EMAIL_SUBJECT = "Relatório Genesys Cloud de Patricia Reis dos Santos"

ATTACHMENT_AGENDAMENTO = "DESEMPENHO DA FILA - AGENDAMENTO ATENTO D1.csv"
ATTACHMENT_AUTORIZACAO = "DESEMPENHO DA FILA - AUTORIZACAO D1.csv"

# ===== DATAS =====

DATA_INICIAL = datetime(2026, 1, 21).date()
DATA_FINAL = datetime.now().date() - timedelta(days=1)

# ===== PASTAS =====

PASTA_AGENDAMENTO = r"\\brsbesrv960\publico\REPORTS\UNIMED BH\HistoricoPython\DesempenhoFilaAgendamentoAtento"

PASTA_AUTORIZACAO = r"\\brsbesrv960\publico\REPORTS\UNIMED BH\HistoricoPython\DesempenhoFilaAutorizacao"

OUTPUT_FINAL = r"\\brsbesrv960\publico\REPORTS\UNIMED BH"

CONSOLIDADO_AGENDAMENTO = "DESEMPENHO DA FILA - AGENDAMENTO ATENTO D1.csv"
CONSOLIDADO_AUTORIZACAO = "DESEMPENHO DA FILA - AUTORIZACAO D1.csv"