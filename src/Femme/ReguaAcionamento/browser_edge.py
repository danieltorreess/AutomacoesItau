from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service


def get_browser(download_dir):
    options = Options()
    options.add_argument("--start-maximized")

    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
    }

    options.add_experimental_option("prefs", prefs)

    service = Service(r"C:\MIS\msedgedriver.exe")

    return webdriver.Edge(service=service, options=options)
