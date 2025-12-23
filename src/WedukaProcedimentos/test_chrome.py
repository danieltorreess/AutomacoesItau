from selenium import webdriver

driver = webdriver.Chrome()
driver.get("https://www.google.com")
input("Chrome abriu. Pressione ENTER para fechar...")
driver.quit()
