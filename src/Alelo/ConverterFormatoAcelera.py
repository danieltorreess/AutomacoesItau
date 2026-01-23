import os
import win32com.client

# Caminho da pasta com os arquivos .xls
pasta = r'C:\Users\ab1541240\Downloads\ACELERA'

# Inicializa o Excel (com configurações para não mostrar janelas)
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
excel.ScreenUpdating = False
excel.Application.WindowState = -4140  # Minimiza a janela do Excel

# Percorre todos os arquivos na pasta
for arquivo in os.listdir(pasta):
    if arquivo.lower().endswith('.xls') and not arquivo.lower().endswith('.xlsx'):
        caminho_completo = os.path.join(pasta, arquivo)
        novo_nome = os.path.splitext(arquivo)[0] + '.xlsx'
        novo_caminho = os.path.join(pasta, novo_nome)

        print(f'Convertendo: {arquivo} -> {novo_nome}')
        try:
            # Abre o arquivo .xls em background
            wb = excel.Workbooks.Open(caminho_completo)

            # Salva como .xlsx (formato 51 é o padrão .xlsx)
            wb.SaveAs(novo_caminho, FileFormat=51)

            # Fecha o arquivo
            wb.Close(SaveChanges=False)

        except Exception as e:
            print(f'Erro ao converter {arquivo}: {e}')

# Encerra o Excel
excel.Quit()
print("Conversão concluída.")


# $origem = "\\brsbesrv960\publico\REPORTS\MRV\MAILING_LOJA_VIRTUAL"
# $destino = "C:\Users\ab1541240\Documents\MRV"
 
# # Cria pasta destino se não existir
# if (!(Test-Path $destino)) { New-Item -ItemType Directory -Path $destino | Out-Null }
 
# # Loop pelos arquivos XLSX
# Get-ChildItem -Path $origem -Filter "Mailing *.xlsx" | ForEach-Object {
#     $arquivoOrigem = $_.FullName
#     $nomeBase = $_.BaseName  # ex: Mailing 07.10.25
#     $dataCurta = $nomeBase.Split(" ")[1] # pega "07.10.25" ou "07.10.2025"
 
#     # Separa as partes do nome
#     $partes = $dataCurta -split "\."
#     $dia = $partes[0]
#     $mes = $partes[1]
#     $ano = $partes[2]
 
#     # Corrige o ano somente se tiver 2 dígitos
#     if ($ano.Length -eq 2) {
#         $ano = "20$ano"
#     }
 
#     $dataFormatada = "$dia.$mes.$ano"
#     $arquivoDestino = Join-Path $destino "Mailing $dataFormatada.csv"
 
#     # Abre Excel COM Object e salva como CSV
#     $excel = New-Object -ComObject Excel.Application
#     $excel.Visible = $false
#     $wb = $excel.Workbooks.Open($arquivoOrigem)
#     $wb.SaveAs($arquivoDestino, 6) # 6 = formato CSV
#     $wb.Close($false)
#     $excel.Quit()
 
#     Write-Host "Convertido: $arquivoDestino"
# }