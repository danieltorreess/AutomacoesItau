import win32com.client as win32


class EmailSender:
    def __init__(self):
        self.outlook = win32.Dispatch("Outlook.Application")

    def enviar_email(self, destinatario: str, assunto: str, corpo: str):
        mail = self.outlook.CreateItem(0)  # MailItem
        mail.To = destinatario
        mail.Subject = assunto
        mail.Body = corpo
        mail.Send()
