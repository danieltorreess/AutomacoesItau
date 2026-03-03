from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time


class ReguaAcionamentoBot:

    def __init__(self, driver, email, password, tenant, config):
        self.driver = driver
        self.wait = WebDriverWait(driver, 30)
        self.email = email
        self.password = password
        self.tenant = tenant
        self.config = config

    # --------------------------------------------------
    # 🔐 Login no sistema
    # --------------------------------------------------
    def login(self):

        self.driver.get(self.config.URL_LOGIN)

        self.wait.until(
            EC.presence_of_element_located((By.ID, "email"))
        ).send_keys(self.email)

        self.driver.find_element(By.ID, "password").send_keys(self.password)
        self.driver.find_element(By.ID, "tenant_name").send_keys(self.tenant)
        self.driver.find_element(By.ID, "submit-btn").click()

    # --------------------------------------------------
    # 📊 Navegar até Relatório → Acionamentos
    # --------------------------------------------------
    def acessar_relatorio(self):

        dropdown = self.wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//a[contains(@class,'dropdown-toggle') and normalize-space()='Relatório']"
            ))
        )

        self.driver.execute_script("""
            arguments[0].click();
            arguments[0].setAttribute('aria-expanded', 'true');
        """, dropdown)

        acionamentos = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//ul[contains(@class,'dropdown-menu')]//a[contains(@href,'/reports/usage')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", acionamentos)

    # --------------------------------------------------
    # 📅 Aplicar filtro de data e exportar CSV
    # --------------------------------------------------
    def aplicar_filtro_e_exportar(self, data_inicio, data_fim):

        inicio_str = data_inicio.strftime("%Y-%m-%dT00:00")
        fim_str = data_fim.strftime("%Y-%m-%dT23:59")

        start_input = self.wait.until(
            EC.presence_of_element_located((By.ID, "startDate-filter"))
        )

        end_input = self.wait.until(
            EC.presence_of_element_located((By.ID, "endDate-filter"))
        )

        # 🔹 Setar datas via JavaScript
        self.driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, start_input, inicio_str)

        self.driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, end_input, fim_str)

        time.sleep(0.5)

        # 🔄 Atualizar relatório
        icone_refresh = self.wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//button[@id='rpt-btn-loader']//i[contains(@class,'glyphicon-refresh')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", icone_refresh)

        time.sleep(3)

        # 📥 Exportar CSV
        export_btn = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[contains(@href,'acionamentos-report')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", export_btn)