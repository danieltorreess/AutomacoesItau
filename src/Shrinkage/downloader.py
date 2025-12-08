import os
from datetime import datetime
from .msg_extractor import MsgExtractor


class ShrinkageDownloader:
    """
    Responsável por:
    - Receber os e-mails encaminhados do Shrinkage.
    - Extrair o .msg interno.
    - Extrair os anexos ATT-YYYYMM de dentro do .msg.
    - Sempre sobrescrever.
    - Garantir que, se houver duplicados no mesmo dia,
      mantenha-se somente o mais recente.
    """

    def __init__(self, output_path):
        self.output_path = output_path
        self.extractor = MsgExtractor()

    def _hora_recebimento(self, email):
        """Retorna o horário do e-mail em formato datetime."""
        return email.ReceivedTime

    def processar_emails(self, emails):
        """
        Para cada e-mail do dia anterior:
        - Abre o .msg interno
        - Extrai anexos ATT-YYYYMM
        - Se houver arquivos duplicados no mesmo dia, mantém o mais recente
        """

        if not emails:
            print("⚠️ Nenhum e-mail encontrado para processar no Shrinkage.")
            return {}

        # Ordena e-mails do mais recente pro mais antigo
        emails = sorted(emails, key=lambda m: m.ReceivedTime, reverse=True)

        arquivos_final = {}  
        # Exemplo:
        # {
        #   "ATT-202512 - HUB Performance.xlsm": "caminho/salvo",
        #   "ATT-202512 - HUB Agentes.xlsm": "caminho/salvo",
        # }

        print(f"📨 Processando {len(emails)} e-mails do dia...")

        for email in emails:
            hora = email.ReceivedTime.strftime("%H:%M:%S")
            print(f"\n🔎 Processando e-mail recebido às {hora}...")

            # Extrai anexos ATT-YYYYMM de dentro do .msg
            anexos = self.extractor.extrair_anexos_att(email, self.output_path)

            if not anexos:
                print("⚠️ Nenhum ATT-YYYYMM encontrado neste e-mail.")
                continue

            # Para cada arquivo extraído, manter somente o mais recente
            for caminho in anexos:
                nome_arquivo = os.path.basename(caminho)

                # Se ainda não temos esse arquivo → usar este
                if nome_arquivo not in arquivos_final:
                    arquivos_final[nome_arquivo] = caminho
                    continue

                # Se já existe, substitui pelo mais recente (tabela já ordenada)
                print(f"♻️ Substituindo versão antiga de {nome_arquivo} pela mais recente.")

                # Nada a fazer porque o arquivo salvo já é sobrescrito sempre,
                # basta atualizar o dicionário
                arquivos_final[nome_arquivo] = caminho

        print("\n✅ Extração concluída!")

        print("\n📁 Arquivos finais considerados para o SHRINKAGE:")
        for nome, path in arquivos_final.items():
            print(f"   → {nome}")

        return arquivos_final
