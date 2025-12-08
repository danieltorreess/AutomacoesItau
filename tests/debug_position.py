import win32com.client as win32

path = r"\\BRSBESRV960\Publico\REPORTS\ITAU\SHRINKAGE\ATT-202512 - HUB Atendimento.xlsm"

excel = win32.Dispatch("Excel.Application")
excel.Visible = True

wb = excel.Workbooks.Open(path)

print("\n=== LISTA DE ABAS E SUAS POSIÇÕES ===")
for i, sh in enumerate(wb.Sheets, start=1):
    print(f"{i}: {sh.Name}")

voz = wb.Sheets("VOZ")
print(f"\nAba VOZ está na posição: {voz.Index}")

linha_total = None
ultima = voz.Cells(voz.Rows.Count, 2).End(-4162).Row

for i in range(1, ultima + 1):
    if str(voz.Cells(i, 2).Value).strip().lower() == "total geral":
        linha_total = i
        break

print(f"Total Geral encontrado em B{linha_total}")

try:
    print("\nExplodindo pivot...")
    voz.Range(f"C{linha_total}").ShowDetail = True
    print("Explodiu sem erro!")
except Exception as e:
    print("ERRO ao explodir pivot:", e)

print("\nAbas após explodir:")
for i, sh in enumerate(wb.Sheets, start=1):
    print(f"{i}: {sh.Name}")

wb.Close(False)
excel.Quit()
