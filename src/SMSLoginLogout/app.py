import zipfile
import shutil
import tempfile
from pathlib import Path
import re


def remover_protecao_excel(
    arquivo_origem: Path,
    arquivo_destino: Path
):
    if not arquivo_origem.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_origem}")

    if arquivo_origem.suffix.lower() != ".xlsx":
        raise ValueError("O arquivo precisa ser .xlsx")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)

        # 1. Extrai o xlsx (zip)
        with zipfile.ZipFile(arquivo_origem, "r") as zip_ref:
            zip_ref.extractall(tmpdir)

        # 2. Remove proteção das planilhas
        sheets_path = tmpdir / "xl" / "worksheets"
        if sheets_path.exists():
            for sheet in sheets_path.glob("*.xml"):
                xml = sheet.read_text(encoding="utf-8")

                # Remove tags <sheetProtection ... />
                xml_limpo = re.sub(
                    r"<sheetProtection[^>]*/>",
                    "",
                    xml,
                    flags=re.IGNORECASE
                )

                sheet.write_text(xml_limpo, encoding="utf-8")

        # 3. Remove proteção da estrutura do workbook
        workbook_xml = tmpdir / "xl" / "workbook.xml"
        if workbook_xml.exists():
            xml = workbook_xml.read_text(encoding="utf-8")

            xml_limpo = re.sub(
                r"<workbookProtection[^>]*/>",
                "",
                xml,
                flags=re.IGNORECASE
            )

            workbook_xml.write_text(xml_limpo, encoding="utf-8")

        # 4. Recompacta para novo xlsx
        with zipfile.ZipFile(arquivo_destino, "w", zipfile.ZIP_DEFLATED) as zip_out:
            for file in tmpdir.rglob("*"):
                if file.is_file():
                    zip_out.write(
                        file,
                        file.relative_to(tmpdir)
                    )


if __name__ == "__main__":
    origem = Path(
        r"C:\Users\ab1541240\OneDrive - ATENTO Global\Documents\ITAU\RelatoriosItau\02 - PowerBI\MIS00000 - SMS_EMAILS\Login e Logout LD.xlsx"
    )

    destino = origem.with_stem(origem.stem + "_SEM_PROTECAO")

    remover_protecao_excel(origem, destino)

    print(f"✔ Arquivo gerado sem proteção:\n{destino}")
