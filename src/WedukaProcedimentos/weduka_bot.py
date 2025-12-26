import time
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

from src.WedukaProcedimentos.utils import (
    get_date_range,
    wait_for_download,
    move_and_rename,
    xlsx_to_csv
)


class WedukaBot:

    def __init__(self, driver, username, password, config):
        self.driver = driver
        self.wait = WebDriverWait(driver, 30)
        self.username = username
        self.password = password
        self.config = config

    def login(self):
        print("[WEDUKA] Acessando página de integração...")
        self.driver.get(self.config.URL_INTEGRATION)

        self.wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Ir para site de autenticação"))
        ).click()

        print("[WEDUKA] Realizando login...")
        self.wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(self.username)
        self.driver.find_element(By.ID, "password").send_keys(self.password)
        self.driver.find_element(By.ID, "btLogin").click()

    def acessar_relatorio(self):
        print("[WEDUKA] Navegando para Relatório de acessos...")
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Procedimentos']"))
        ).click()

        self.wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Relatório de acessos"))
        ).click()

    def extrair_repositorio(self, nome_repo):
        print(f"[WEDUKA] Iniciando extração do repositório: {nome_repo}")

        # ------------------------------------------------------------------
        # Dropdown de repositório
        # ------------------------------------------------------------------
        dropdown_btn = self.wait.until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, "button[data-id='ID_Document_Repository']")
            )
        )
        dropdown_btn.click()

        # Input de busca do dropdown (Bootstrap Select)
        search_input = self.wait.until(
            EC.visibility_of_element_located(
                (By.CSS_SELECTOR, ".dropdown-menu.show .bs-searchbox input")
            )
        )

        # 🔒 Blindagem contra ElementNotInteractableException
        self.driver.execute_script(
            "arguments[0].scrollIntoView(true);", search_input
        )
        self.driver.execute_script(
            "arguments[0].focus();", search_input
        )
        self.driver.execute_script(
            "arguments[0].value = '';", search_input
        )
        time.sleep(0.3)

        search_input.send_keys(nome_repo)

        # Seleciona a opção correta
        option = self.wait.until(
            EC.visibility_of_element_located(
                (By.XPATH, f"//span[@class='text' and normalize-space()='{nome_repo}']")
            )
        )

        self.safe_click(option)
        print(f"[WEDUKA] Repositório selecionado: {nome_repo}")

        # ------------------------------------------------------------------
        # Período
        # ------------------------------------------------------------------
        periodo = get_date_range()
        date_input = self.driver.find_element(By.ID, "DateRange")
        date_input.clear()
        date_input.send_keys(periodo)
        date_input.send_keys(Keys.ENTER)

        print(f"[WEDUKA] Período aplicado: {periodo}")

        # ------------------------------------------------------------------
        # Pesquisar
        # ------------------------------------------------------------------
        btn_pesquisar = self.wait.until(
            EC.element_to_be_clickable((By.ID, "submitFilter"))
        )
        self.safe_click(btn_pesquisar)

        print("[WEDUKA] Executando pesquisa...")

        # ------------------------------------------------------------------
        # Exportação (> 500 linhas)
        # ------------------------------------------------------------------
        export_link = self.wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//a[contains(@href,'export=True')]")
            )
        )

        export_url = export_link.get_attribute("href")
        print("[WEDUKA] Resultado acima de 500 linhas — iniciando download...")

        self.driver.get(export_url)

        file_xlsx = wait_for_download(self.config.DOWNLOAD_DIR)

        base_name = f"{self.config.FILE_PREFIX}{nome_repo.replace(' ', '').lower()}"
        csv_name = f"{base_name}.csv"
        csv_path = self.config.DEST_DIR / csv_name

        # Converte para CSV direto na pasta final
        xlsx_to_csv(file_xlsx, csv_path)

        # Remove o XLSX temporário
        file_xlsx.unlink(missing_ok=True)

        print(f"[WEDUKA] Arquivo CSV salvo com sucesso em: {csv_path}")
        print("-" * 70)

    def safe_click(self, element):
        self.driver.execute_script(
            "arguments[0].scrollIntoView({block: 'center'});", element
        )
        try:
            element.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", element)
