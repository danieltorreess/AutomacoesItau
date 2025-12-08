import win32com.client as win32
import time

path = r"\\BRSBESRV960\Publico\REPORTS\ITAU\SHRINKAGE\ATT-202512 - HUB Atendimento.xlsm"

excel = win32.Dispatch("Excel.Application")
excel.Visible = True
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(path)

sheet = wb.Sheets("VOZ")

# Localizar Total Geral na coluna B
linha_total = None
ultima = sheet.Cells(sheet.Rows.Count, 2).End(-4162).Row

for i in range(1, ultima + 1):
    if str(sheet.Cells(i, 2).Value).strip().lower() == "total geral":
        linha_total = i
        break

print(f"Total Geral em B{linha_total}")

print("Explodindo pivot...")
sheet.Range(f"C{linha_total}").ShowDetail = True

print("\n\n=== PAUSA ===")
print("Uma aba 'DetalhesX' foi criada.")
print("👉 Abra ela manualmente e CONFIRME:")
print("- Há dados?")
print("- Os dados começam em qual linha?")
print("- Qual a primeira coluna com dados?")
print("- Qual a última coluna com dados?")
print("\nFeche o Excel quando terminar de inspecionar.")

while True:
    try:
        time.sleep(1)
    except KeyboardInterrupt:
        break

wb.Close(False)
excel.Quit()
