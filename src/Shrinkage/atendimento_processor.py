import win32com.client as win32
import pandas as pd
import os
import tempfile


class AtendimentoProcessor:

    def __init__(self):
        self.excel = win32.Dispatch("Excel.Application")
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.AskToUpdateLinks = False
        self.excel.AlertBeforeOverwriting = False
        self.excel.AutoRecover.Enabled = False
        self.excel.FeatureInstall = 0
        self.excel.CutCopyMode = False

    # --------------------------------------------------------
    def _abrir_excel(self, caminho):
        try:
            return self.excel.Workbooks.Open(caminho)
        except Exception as e:
            print(f"❌ Erro ao abrir {caminho}: {e}")
            return None

    # --------------------------------------------------------
    def _localizar_total_geral(self, sheet):
        ultima = sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row
        for i in range(1, ultima + 1):
            if str(sheet.Cells(i, 2).Value).strip().lower() == "total geral":
                return i
        return None

    # --------------------------------------------------------
    def _explodir_para_csv(self, sheet, linha_total):

        indice_original = sheet.Index
        sheet.Range(f"C{linha_total}").ShowDetail = True

        detalhes_sheet = sheet.Parent.Sheets(indice_original)

        temp_path = os.path.join(tempfile.gettempdir(), f"det_{indice_original}.csv")
        if os.path.exists(temp_path):
            os.remove(temp_path)

        detalhes_sheet.SaveAs(temp_path, FileFormat=6)
        return temp_path

    # --------------------------------------------------------
    def _ler_csv(self, caminho_csv):
        try:
            df = pd.read_csv(caminho_csv, encoding="latin1", sep=",")
            df.columns = df.columns.str.strip()
            return df
        except Exception as e:
            print(f"❌ Falha ao ler CSV: {e}")
            return pd.DataFrame()

    # --------------------------------------------------------
    def _montar_espelho(self, df_voz, df_dig, caminho_saida):
        """
        ✔ Estrutura do espelho = estrutura da tabela dinâmica
        ✔ Cabeçalho vem da VOZ
        ✔ DIGITAL entra sem cabeçalho
        """

        # Remover cabeçalho da DIGITAL
        df_dig_sem_header = df_dig.copy()
        df_dig_sem_header.columns = df_voz.columns  # garante alinhamento
        df_dig_sem_header = df_dig_sem_header.iloc[1:]  # remove a linha de header real

        df_final = pd.concat([df_voz, df_dig_sem_header], ignore_index=True)

        df_final.to_excel(caminho_saida, index=False, sheet_name="VOZ_DIGITAL")

        print(f"📁 Espelho criado em: {caminho_saida}")

    # --------------------------------------------------------
    def processar(self, caminho_att, caminho_espelho_saida):

        wb = self._abrir_excel(caminho_att)
        if wb is None:
            return False

        print("\n🔎 Processando VOZ...")
        aba_voz = wb.Sheets("VOZ")
        linha_voz = self._localizar_total_geral(aba_voz)
        df_voz = self._ler_csv(self._explodir_para_csv(aba_voz, linha_voz))

        print("\n🔎 Processando DIGITAL...")
        aba_dig = wb.Sheets("DIGITAL")
        linha_dig = self._localizar_total_geral(aba_dig)
        df_dig = self._ler_csv(self._explodir_para_csv(aba_dig, linha_dig))

        wb.Close(False)

        # Criar o espelho final
        self._montar_espelho(df_voz, df_dig, caminho_espelho_saida)

        self.excel.Quit()
        return True
