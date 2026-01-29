import os
import win32com.client as win32
class MsgExtractor:
    """
    Extrai anexos ATT-YYYYMM diretamente do e-mail principal.
    """

    def extrair_anexos_att(self, email, output_path):
        os.makedirs(output_path, exist_ok=True)

        encontrados = []

        for att in email.Attachments:
            nome = att.FileName.lower()

            if not nome.startswith("att-"):
                continue

            if not (
                nome.endswith(".xlsm")
                or nome.endswith(".xlsx")
                or nome.endswith(".xlsb")
            ):
                continue

            destino = os.path.join(output_path, att.FileName)

            try:
                att.SaveAsFile(destino)
                print(f"📁 Anexo salvo: {destino}")
                encontrados.append(destino)
            except Exception as e:
                print(f"❌ Erro ao salvar anexo {att.FileName}: {e}")

        if not encontrados:
            print("⚠️ Nenhum arquivo ATT-* encontrado no e-mail.")

        return encontrados

# class MsgExtractor:
#     """
#     Responsável por abrir o anexo .msg do e-mail encaminhado
#     e extrair de dentro dele os anexos reais (ATT-YYYYMM).
#     """

#     def __init__(self):
#         self.outlook = win32.Dispatch("Outlook.Application").GetNamespace("MAPI")

#     def _abrir_msg_interno(self, caminho_msg_temp):
#         """
#         Abre um arquivo .msg temporário e retorna o MailItem interno.
#         """
#         return self.outlook.OpenSharedItem(caminho_msg_temp)

#     def extrair_anexos_att(self, email_encaminhado, output_path):
#         """
#         1. Extrai o .msg interno do email encaminhado.
#         2. Abre esse .msg como MailItem.
#         3. Extrai os anexos ATT-YYYYMM.
#         4. Salva tudo sobrescrevendo na pasta destino.
#         """

#         # Criar pasta destino caso não exista
#         os.makedirs(output_path, exist_ok=True)

#         # Procurar o .msg dentro do e-mail encaminhado
#         msg_anexo = None
#         for att in email_encaminhado.Attachments:
#             if att.FileName.lower().endswith(".msg"):
#                 msg_anexo = att
#                 break

#         if msg_anexo is None:
#             print("⚠️ Nenhum .msg encontrado dentro do e-mail encaminhado.")
#             return []

#         # Salvar o .msg TEMPORARIAMENTE
#         temp_path = os.path.join(output_path, "_temp_email.msg")
#         try:
#             msg_anexo.SaveAsFile(temp_path)
#         except Exception as e:
#             print(f"❌ Erro ao salvar .msg temporário: {e}")
#             return []

#         # Abrir o .msg salvo como MailItem interno
#         try:
#             email_original = self._abrir_msg_interno(temp_path)
#         except Exception as e:
#             print(f"❌ Erro ao abrir o .msg interno: {e}")
#             return []

#         print(f"📨 E-mail original aberto: {email_original.Subject}")

#         # Procurar anexos que começam com ATT-YYYYMM
#         encontrados = []

#         for att in email_original.Attachments:
#             nome = att.FileName

#             # Aceitar arquivos Excel que começam com ATT-YYYYMM
#             if not nome.lower().startswith("att-"):
#                 continue

#             if not (nome.lower().endswith(".xlsm")
#                     or nome.lower().endswith(".xlsx")
#                     or nome.lower().endswith(".xlsb")):
#                 continue

#             destino = os.path.join(output_path, nome)

#             try:
#                 att.SaveAsFile(destino)
#                 print(f"📁 Anexo salvo: {destino}")
#                 encontrados.append(destino)
#             except Exception as e:
#                 print(f"❌ Erro ao salvar anexo {nome}: {e}")

#         # Remover arquivo temporário
#         try:
#             os.remove(temp_path)
#         except:
#             pass

#         if not encontrados:
#             print("⚠️ Nenhum arquivo ATT- encontrado dentro do e-mail original.")

#         return encontrados
