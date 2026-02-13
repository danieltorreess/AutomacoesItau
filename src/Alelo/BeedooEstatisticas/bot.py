from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time


class BeedooEstatisticasBot:

    def __init__(self, driver, email, password, config):
        self.driver = driver
        self.wait = WebDriverWait(driver, 30)
        self.email = email
        self.password = password
        self.config = config

    # --------------------------------------------------
    def login(self):
        self.driver.get(self.config.URL_LOGIN)

        self.wait.until(
            EC.presence_of_element_located((By.ID, "inputLogin"))
        ).send_keys(self.email)

        self.driver.find_element(By.ID, "inputPassword").send_keys(self.password)

        # botão Login (fica habilitado após preencher)
        botao = self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(.,'Login')]"))
        )

        botao.click()

    # --------------------------------------------------
    def abrir_estatisticas(self):

        # 🔹 Espera página carregar após login
        time.sleep(5)

        # 🔹 Clicar no ícone grid (menu principal)
        menu = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//i[contains(@class,'fa-th')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", menu)

        # 🔹 Clicar em Estatísticas
        estatisticas = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[contains(@href,'/adm/report')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", estatisticas)

    # --------------------------------------------------
    def abrir_feed_analitico(self):

        # 🔹 Espera página estatísticas carregar
        time.sleep(5)

        # 🔹 Clica no "+" do Feed
        botao_mais = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//div[contains(@class,'more-stat') and text()='+']"
            ))
        )

        self.driver.execute_script("arguments[0].click();", botao_mais)

        # 🔹 Clica em Filtrar
        filtrar = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//a[contains(@class,'open-filter') and contains(.,'Filtrar')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", filtrar)

        # 🔹 Clica no filtro de período
        periodo = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//span[contains(.,'Últimos 30 dias')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", periodo)

        # 🔹 Seleciona "Últimos 7 dias"
        ultimos7 = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//li[contains(.,'Últimos 7 dias')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", ultimos7)

        time.sleep(1)

        # 🔹 Clicar em Relatório Analítico
        relatorio = self.wait.until(
            EC.element_to_be_clickable((
                By.XPATH,
                "//span[contains(@title,'Exportar relatório')]"
            ))
        )

        self.driver.execute_script("arguments[0].click();", relatorio)
