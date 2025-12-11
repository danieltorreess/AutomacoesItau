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
        Nova lógica:
        - Sempre começar pelo e-mail mais recente
        - Extrair anexos ATT-* desse e-mail
        - Anotar quais bases já vieram
        - Se estiver faltando alguma base, continuar indo para os próximos e-mails
        """

        if not emails:
            print("⚠️ Nenhum e-mail encontrado para processar.")
            return {}

        # Ordenar: mais recente primeiro
        emails = sorted(emails, key=lambda m: m.ReceivedTime, reverse=True)

        print(f"📨 Encontrados {len(emails)} e-mails. Iniciando processamento inteligente...\n")

        # Bases esperadas
        bases_esperadas = {
            "agentes": None,
            "atendimento": None,
            "performance": None
        }

        # Controle final
        arquivos_final = {}

        # Para cada e-mail (apenas se faltar base)
        for email in emails:

            faltando = [k for k, v in bases_esperadas.items() if v is None]
            if not faltando:
                break  # já temos tudo

            hora = email.ReceivedTime.strftime("%H:%M:%S")
            print(f"🔎 Processando e-mail das {hora}...")

            anexos = self.extractor.extrair_anexos_att(email, self.output_path)

            for caminho in anexos:
                nome = os.path.basename(caminho).lower()

                if "agentes" in nome and bases_esperadas["agentes"] is None:
                    bases_esperadas["agentes"] = caminho
                    arquivos_final[os.path.basename(caminho)] = caminho

                if "atendimento" in nome and bases_esperadas["atendimento"] is None:
                    bases_esperadas["atendimento"] = caminho
                    arquivos_final[os.path.basename(caminho)] = caminho

                if "performance" in nome and bases_esperadas["performance"] is None:
                    bases_esperadas["performance"] = caminho
                    arquivos_final[os.path.basename(caminho)] = caminho

        print("\n📌 Resultado final das bases encontradas:")

        for base, path in bases_esperadas.items():
            if path:
                print(f"   ✓ {base.capitalize()} encontrado: {path}")
            else:
                print(f"   ⚠️ {base.capitalize()} NÃO encontrado nos últimos e-mails.")

        print("\n✅ Seleção final concluída!\n")

        return arquivos_final
