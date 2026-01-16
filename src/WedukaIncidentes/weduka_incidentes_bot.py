from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from src.WedukaIncidentes.utils import (
    get_days_of_month_until_yesterday,
    format_day_range,
    wait_for_download,
    xlsx_to_csv
)


class WedukaIncidentesBot:

    def __init__(self, driver, username, password, config):
        self.driver = driver
        self.wait = WebDriverWait(driver, 30)
        self.username = username
        self.password = password
        self.config = config

    def login(self):
        print("[INCIDENTES] Login no Weduka")
        self.driver.get(self.config.URL_INTEGRATION)

        # self.wait.until(
        #     EC.element_to_be_clickable((By.LINK_TEXT, "Ir para site de autenticação"))
        # ).click()

        self.wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//a[contains(text(),'autenticação')]")
            )
        ).click()

        self.wait.until(EC.presence_of_element_located((By.ID, "username"))).send_keys(self.username)
        self.driver.find_element(By.ID, "password").send_keys(self.password)
        self.driver.find_element(By.ID, "btLogin").click()

    def acessar_relatorio(self):
        print("[INCIDENTES] Acessando relatório de incidentes")

        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[text()='Incidentes']"))
        ).click()

        self.wait.until(
            EC.element_to_be_clickable((By.LINK_TEXT, "Relatório"))
        ).click()

        # Agrupar por pessoas
        self.wait.until(
            EC.element_to_be_clickable((By.XPATH, "//label[@for='GroupPeople']"))
        ).click()

    def extrair_por_dia(self):
        for day in get_days_of_month_until_yesterday():
            periodo = format_day_range(day)
            print(f"[INCIDENTES] Processando dia: {periodo}")

            date_input = self.wait.until(
                EC.element_to_be_clickable((By.ID, "DateRange"))
            )
            date_input.clear()
            date_input.send_keys(periodo)
            date_input.send_keys(Keys.ENTER)

            # 👇 PESQUISAR (JS CLICK)
            btn_pesquisar = self.wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@type='submit' and @value='Pesquisar']")
                )
            )
            self.driver.execute_script("arguments[0].click();", btn_pesquisar)

            # Aguarda possível link de exportação
            try:
                export_link = self.wait.until(
                    EC.presence_of_element_located(
                        (By.XPATH, "//a[contains(@href,'export=True')]")
                    )
                )
            except Exception:
                print(f"[INCIDENTES] Sem base para {periodo} (ou >100k linhas)")
                continue

            export_url = export_link.get_attribute("href")
            print(f"[INCIDENTES] Exportando base do dia {day.strftime('%d/%m/%Y')}")

            self.driver.get(export_url)

            file_xlsx = wait_for_download(self.config.DOWNLOAD_DIR)

            file_name = f"{self.config.FILE_PREFIX}{day.strftime('%d%m%Y')}.csv"
            csv_path = self.config.DEST_DIR / file_name

            xlsx_to_csv(file_xlsx, csv_path)
            file_xlsx.unlink(missing_ok=True)

            print(f"[INCIDENTES] Arquivo salvo: {csv_path}")
            print("-" * 60)
