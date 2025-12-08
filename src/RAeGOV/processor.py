import os
import win32com.client as win32


OUTPUT_PATH = r"\\BRSBESRV960\publico\REPORTS\001 CARGAS\001 EXCEL\ITAU\QUALITATIVO"


class RaeGovProcessor:

    def __init__(self):
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False

    def salvar_ra(self, attachment):
        destino = os.path.join(OUTPUT_PATH, attachment.FileName)
        attachment.SaveAsFile(destino)
        print(f"📁 RA salvo: {destino}")

    def salvar_gov(self, attachment):
        destino = os.path.join(OUTPUT_PATH, "Checklist - Canais Criticos Consignado - GOV.xlsx")
        caminho_tmp = destino + ".tmp"

        # salvar primeiro temporário
        attachment.SaveAsFile(caminho_tmp)

        # abrir para renomear aba
        wb = self.excel.Workbooks.Open(caminho_tmp)

        alvo = None
        for sh in wb.Sheets:
            nome = sh.Name.lower().replace("á", "a").replace("ã", "a")
            if "analise" in nome and "validador" in nome:
                alvo = sh
                break

        if alvo:
            alvo.Name = "Analise sem validador"
            print("📝 Aba renomeada para 'Analise sem validador'")
        else:
            print("⚠️ Nenhuma aba correspondente encontrada no GOV.")

        wb.SaveAs(destino)
        wb.Close(False)

        os.remove(caminho_tmp)

        print(f"📁 GOV salvo: {destino}")

    def processar_anexos(self, anexos):
        for att in anexos:
            nome = att.FileName.lower()

            if "ra" in nome or "log" in nome:
                self.salvar_ra(att)

            elif "gov" in nome:
                self.salvar_gov(att)

            else:
                print(f"⚠️ Anexo ignorado: {att.FileName}")
