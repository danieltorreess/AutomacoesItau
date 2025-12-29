# src/ItauSiteGov/navigation.py
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

from .utils import wait_for_element, wait_and_click
from .config import FORNECEDORES, URL_RELATORIO_GERENCIAL

def select_fornecedor(driver, fornecedor_key: str):
    """
    Seleciona o fornecedor no combo 'Qual fornecedor deseja gerenciar?'
    """
    if fornecedor_key not in FORNECEDORES:
        raise ValueError(f"Fornecedor inválido: {fornecedor_key}")

    fornecedor = FORNECEDORES[fornecedor_key]

    print(f"🏦 Selecionando fornecedor: {fornecedor['label']}")

    select_element = wait_for_element(driver, By.ID, "menu_sel_fornecedor")
    select = Select(select_element)

    select.select_by_value(fornecedor["value"])

def open_relatorios_gerenciais(driver):
    """
    Navega pelo menu até Relatórios Gerenciais
    """
    print("📊 Abrindo Relatórios Gerenciais...")

    # Abrir dropdown Menu
    wait_and_click(
        driver,
        By.XPATH,
        "//button[contains(@class,'dropdown-toggle') and contains(.,'Menu')]"
    )

    # Clicar em Relatórios Gerenciais
    wait_and_click(
        driver,
        By.XPATH,
        "//a[contains(@href,'/pages/relatorio/gerencial/abrir')]"
    )
