import re
import shutil
from pathlib import Path
import pandas as pd


# -------------------------------
# Diretório de origem (Downloads)
# -------------------------------
DOWNLOADS = Path.home() / "Downloads"


# -------------------------------
# CONFIGURAÇÃO DAS BASES
# -------------------------------
BASES = [
    {
        "nome": "NPS CHAT",
        "regex": re.compile(r"DASH NPS \d{6}\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\NPS CHAT CARTÕES\Extracao Python"),
        "tipo": "csv"
    },
    {
        "nome": "BASE CSAT",
        "regex": re.compile(r"BASE CSAT \d{6}\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\NPS VOZ CARTOES\Extracao Python"),
        "tipo": "csv"
    },
    {
        "nome": "PERFIL QUALIDADE",
        "regex": re.compile(r"^\d{6}_Perfil\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\PERFIL DE QUALIDADE CARTÕES\Extracao Python"),
        "tipo": "csv"
    },
    {
        "nome": "RETENÇÃO CARTÕES",
        "regex": re.compile(r"RETENÇÃO CARTÕES \d{6}\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\RETENCAO_CARTOES\Extracao Python"),
        "tipo": "csv_para_xlsx"
    },
    {
        "nome": "DIAGNÓSTICO NPS_CSAT",
        "regex": re.compile(r"DIAGNÓSTICO NPS_CSAT\.xlsb$", re.IGNORECASE),
        "destino": Path(r"\\brsbesrv960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\DIAGNOSTICOs\NPS_CSAT"),
        "tipo": "xlsb"
    }
]


# -------------------------------
# Converter CSV para XLSX
# -------------------------------
def converter_csv_para_xlsx(caminho_csv: Path) -> Path:
    print("   🔄 Convertendo CSV para XLSX (Excel-friendly)...")

    df = pd.read_csv(
        caminho_csv,
        sep=";",
        dtype=str,
        encoding="utf-8"
    )

    caminho_xlsx = caminho_csv.with_suffix(".xlsx")

    df.to_excel(
        caminho_xlsx,
        index=False,
        engine="openpyxl"
    )

    print(f"   ✔ XLSX gerado: {caminho_xlsx.name}")
    return caminho_xlsx


# -------------------------------
# Função principal
# -------------------------------
def mover_arquivos():
    print("\n===============================")
    print(" PROCESSANDO BASES ITAU SCOUT ")
    print("===============================\n")

    for base in BASES:
        print(f"\n➡ Procurando arquivos para: {base['nome']}")

        arquivos = [
            f for f in DOWNLOADS.iterdir()
            if f.is_file() and base["regex"].match(f.name)
        ]

        if not arquivos:
            print("   ❌ Nenhum arquivo encontrado.")
            continue

        print(f"   ✔ {len(arquivos)} arquivo(s) encontrado(s):")
        for f in arquivos:
            print(f"     - {f.name}")

        print("\n   ➡ Movendo arquivos...")

        for f in arquivos:
            destino_final = base["destino"] / f.name
            try:
                shutil.move(str(f), str(destino_final))
                print(f"     ✔ Movido: {f.name}")

                # 🔥 TRATAMENTO ESPECIAL: RETENÇÃO CARTÕES
                if base["tipo"] == "csv_para_xlsx":
                    caminho_xlsx = converter_csv_para_xlsx(destino_final)

                    # (opcional) apagar o CSV original
                    # destino_final.unlink()
                    # print("   🗑 CSV original removido.")

            except Exception as e:
                print(f"     ❌ Erro ao mover {f.name}: {e}")

    print("\n✔ Finalizado.\n")


# -------------------------------
# Execução
# -------------------------------
if __name__ == "__main__":
    mover_arquivos()
