from .email_sender import EmailSender

DESTINATARIO = "svc_kpi_mis@atento.com"
CORPO_EMAIL = "Relatório enviado."

ASSUNTOS = [
    "MIS1191 - MICRO-GESTÃO_ATENDIMENTO_RETENÇÃO FLEX",
    "MIS30986 - TABULAÇÃO_BACKOFFICE_FRAUDE",
    "MIS30455 - TABULAÇÃO_BACKOFFICE_CSR IC",
    "MIS02266 - PROVIMENTO_ATENDIMENTO_PA´S LOGADAS",
    "MIS30403 - INTRADAY_RECEPTIVO_CLICKPAG",
    "MIS30744 - BASES_ANALITICAS_ATENDIMENTO",
    "MIS30930 - INTRADAY_ATENDIMENTO_SLA SUPORTE E-MAIL",
    "MIS31144 - INTRADAY_CALLBACK",
    "MIS90132 - INTRADAY_ATENDIMENTO_ ATIVO WELCOME CALL FINANCIAMENTO",
    "MIS1319 - INTRADAY_RECEPTIVO_RELACIONAMENTO TEF",
    "MIS31018 - INTRADAY BACKOFFICE SEGUNDO NIVEL PJ",
    "MIS30758 - INTRADAY_BACKOFFICE_2º NÍVEL PÓS VENDA",
    "MIS31414 - Painel de Volumetria Intraday",
    "MIS31384 - INTRADAY - ANALISE PAGAMENTO",
    "MIS31449 - INTRADAY_ITAU_NCG_FALHAS_OPERACIONAIS",
    "MIS31022 - Gerencial Canais Criticos - RA Consignado",
    "MIS31209 - Gerencial Canais Criticos - Consumidor.Gov",
    "MIS31210 - GERENCIAL CANAIS CRITICOS - QUALITY",
    "MIS90128 - Gerencial_Canais Criticos_RA Itaucred",
    "MIS31182 - INTRA-HORA_ATENDIMENTO_FDI",
    "MIS31381 - One-Page Multi Centrais",
    "MIS31478 - Treinamento Weduka",
    "MIS31422 - Micro Gestão - Atendimento - Multi Centrais",
    "MIS31457 - MICRO_GESTÃO_QUALIDADE",
    "MIS31174 - BASES-ANALITICAS_ATENDIMENTO_CHARGEBACK_3547",
    "MIS31175 - BASES-ANALITICAS_ATENDIMENTO_CHARGEBACK_3580",
    "MIS31176 - BASES-ANALITICAS_ATENDIMENTO_CHARGEBACK_FEITO_CONFERIDO",
    "MIS31180 - BASES-ANALITICAS_ATENDIMENTO_CHARGEBACK_DAP",
]


def main():
    print("\n🚀 Iniciando envio de relatórios...")

    sender = EmailSender()

    for assunto in ASSUNTOS:
        print(f"📧 Enviando e-mail: {assunto}")
        sender.enviar_email(
            destinatario=DESTINATARIO,
            assunto=assunto,
            corpo=CORPO_EMAIL
        )

    print("\n🎉 Todos os e-mails foram enviados com sucesso!")


if __name__ == "__main__":
    main()
