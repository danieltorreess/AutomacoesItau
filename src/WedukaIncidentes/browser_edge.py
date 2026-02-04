from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service


def get_browser(download_dir):
    edge_options = Options()
    edge_options.add_argument("--start-maximized")

    prefs = {
        "download.default_directory": str(download_dir),
        "download.prompt_for_download": False,
        "directory_upgrade": True,
    }
    edge_options.add_experimental_option("prefs", prefs)

    service = Service(
        executable_path=r"C:\MIS\msedgedriver.exe"
    )

    driver = webdriver.Edge(
        service=service,
        options=edge_options
    )

    return driver
