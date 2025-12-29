# src/ItauSiteGov/login.py
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from .config import GOVBR_CPF, GOVBR_PASSWORD, URL_LOGIN
from .utils import wait_and_click


def login_govbr(driver):
    wait = WebDriverWait(driver, 60)

    print("🔐 Abrindo página de login gov.br...")
    driver.get(URL_LOGIN)

    print("🖱️ Clicando em 'Entrar com gov.br'...")
    wait.until(
        EC.element_to_be_clickable(
            (By.XPATH, "//a[contains(@class,'gestor') and contains(.,'gov.br')]")
        )
    ).click()

    print("⏳ Aguardando campo de CPF...")
    cpf_input = wait.until(
        EC.visibility_of_element_located((By.ID, "accountId"))
    )
    cpf_input.clear()
    cpf_input.send_keys(GOVBR_CPF)

    wait.until(
        EC.element_to_be_clickable((By.ID, "enter-account-id"))
    ).click()

    print("⏳ Aguardando campo de senha...")
    password_input = wait.until(
        EC.visibility_of_element_located((By.ID, "password"))
    )
    password_input.clear()
    password_input.send_keys(GOVBR_PASSWORD)

    wait.until(
        EC.element_to_be_clickable((By.ID, "submit-button"))
    ).click()

    print("⏳ Aguardando conclusão do login...")
    wait.until(
        EC.presence_of_element_located((By.ID, "menu_sel_fornecedor"))
    )

    print("✅ Login gov.br concluído com sucesso")
