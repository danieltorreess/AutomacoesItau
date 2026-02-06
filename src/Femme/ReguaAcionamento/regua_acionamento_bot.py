from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
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
    def login(self):
        self.driver.get(self.config.URL_LOGIN)

        self.wait.until(EC.presence_of_element_located((By.ID, "email"))).send_keys(self.email)
        self.driver.find_element(By.ID, "password").send_keys(self.password)
        self.driver.find_element(By.ID, "tenant_name").send_keys(self.tenant)

        self.driver.find_element(By.ID, "submit-btn").click()

    # --------------------------------------------------
    def acessar_relatorio(self):
        # 1️⃣ Força abertura do dropdown "Relatório"
        dropdown = self.wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//a[contains(@class,'dropdown-toggle') and normalize-space()='Relatório']"
            ))
        )

        # Força o Bootstrap a abrir o menu
        self.driver.execute_script("""
            arguments[0].click();
            arguments[0].setAttribute('aria-expanded', 'true');
        """, dropdown)

        # 2️⃣ Aguarda o item "Acionamentos" DENTRO do dropdown
        acionamentos = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//ul[contains(@class,'dropdown-menu')]//a[contains(@href,'/reports/usage')]"
            ))
        )

        # 3️⃣ Clique real no item correto
        self.driver.execute_script("arguments[0].click();", acionamentos)

    # --------------------------------------------------
    def aplicar_filtro_e_exportar(self):

        hoje = datetime.today()
        data_inicio = hoje - timedelta(days=5)
        data_fim = hoje - timedelta(days=1)

        inicio_str = data_inicio.strftime("%Y-%m-%dT00:00")
        fim_str = data_fim.strftime("%Y-%m-%dT23:59")

        start_input = self.wait.until(
            EC.presence_of_element_located((By.ID, "startDate-filter"))
        )
        end_input = self.wait.until(
            EC.presence_of_element_located((By.ID, "endDate-filter"))
        )

        # ==================================================
        # 1️⃣ SETA DATAS VIA JS (obrigatório nesse sistema)
        # ==================================================
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
        # ==================================================
        # 2️⃣ CLICA NO ÍCONE DE ATUALIZAR
        # ==================================================
        icone_refresh = self.wait.until(
            EC.presence_of_element_located((
                By.XPATH,
                "//button[@id='rpt-btn-loader']//i[contains(@class,'glyphicon-refresh')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", icone_refresh)

        # ==================================================
        # 3️⃣ AGUARDA PROCESSAMENTO DO RELATÓRIO
        # ==================================================
        time.sleep(3)  # suficiente para esse sistema
        
        # ==================================================
        # 4️⃣ EXPORTAR CSV
        # ==================================================
        export_btn = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[contains(@href,'acionamentos-report')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", export_btn)
