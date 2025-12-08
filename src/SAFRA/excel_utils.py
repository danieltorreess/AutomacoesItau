import win32com.client as win32


def desocultar_abas(caminho_arquivo):
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(caminho_arquivo)

    for sheet in wb.Sheets:
        try:
            sheet.Visible = True
        except:
            pass

    wb.Save()
    wb.Close()
    excel.Quit()
