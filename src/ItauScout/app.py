import re
import shutil
from pathlib import Path


# -------------------------------
# Diretório de origem (Downloads)
# -------------------------------
DOWNLOADS = Path.home() / "Downloads"



# -------------------------------
# CONFIGURAÇÃO DAS 5 BASES
# -------------------------------
BASES = [
    {
        "nome": "NPS CHAT",
        "regex": re.compile(r"DASH NPS \d{6}\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\NPS CHAT CARTÕES\Extracao Python")
    },
    {
        "nome": "BASE CSAT",
        "regex": re.compile(r"BASE CSAT \d{6}\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\NPS VOZ CARTOES\Extracao Python")
    },
    {
        "nome": "PERFIL QUALIDADE",
        "regex": re.compile(r"^\d{6}_Perfil\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\PERFIL DE QUALIDADE CARTÕES\Extracao Python")
    },
    {
        "nome": "RETENÇÃO CARTÕES",
        "regex": re.compile(r"RETENÇÃO CARTÕES \d{6}\.csv$", re.IGNORECASE),
        "destino": Path(r"\\BRSBESRV960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\RETENCAO_CARTOES\Extracao Python")
    },
    {
        "nome": "DIAGNÓSTICO NPS_CSAT",
        "regex": re.compile(r"DIAGNÓSTICO NPS_CSAT\.xlsb$", re.IGNORECASE),
        "destino": Path(r"\\brsbesrv960\Publico\REPORTS\ITAU\ITAU_SCOUT_NPS\BASES ORIGINAIS\DIAGNOSTICOs\NPS_CSAT")
    }
]



# -------------------------------
# Função para mover arquivos
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
            except Exception as e:
                print(f"     ❌ Erro ao mover {f.name}: {e}")

    print("\n✔ Finalizado.\n")



# -------------------------------
# Execução
# -------------------------------
if __name__ == "__main__":
    mover_arquivos()
