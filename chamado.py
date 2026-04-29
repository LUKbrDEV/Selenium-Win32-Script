import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

Options = Options()
Options.debugger_address = "127.0.0.1:9222"
# Carregar dados
df = pd.read_csv("C:\chamados.csv")

driver = webdriver.Chrome()
driver = webdriver.Chrome(options=Options)
driver.get("https://intranet.sicoobcredicopa.com.br/chamados/abrirchamado")

time.sleep(5) # Esperar a página carregar

for _, row in df.iterrows():
    # Preencher campos
    select = Select(driver.find_element(By.NAME, "agencia"))
    select.select_by_visible_text("PA99 - Centro Administrativo")
    select = Select(driver.find_element(By.NAME, "setor"))
    select.select_by_visible_text("Suporte Organizacional")
    time.sleep(5) # Esperar o dropdown carregar
    driver.find_element(By.LINK_TEXT, "Pagamentos 4119").click()
    driver.find_element(By.NAME, "Pagamentos 4119").click()
    driver.find_element(By.NAME, "Fornecedor").send_keys(row["Fornecedor"])
    driver.find_element(By.NAME, "CPF_CNPJ").send_keys(str(row["CPF_CNPJ"]))
    driver.find_element(By.NAME, "Valor").send_keys(str(row["Valor"]))
    driver.find_element(By.NAME, "DataVencimento").send_keys(str(row["Vencimento"]))
    driver.find_element(By.NAME, "FormaPagamento").send_keys(row["FormaPagamento"])


    driver.find_element(By.NAME, "DescricaoPagamento").send_keys(row["Descricao"])

    # Anexar PDF
    driver.find_element(By.NAME, "inputFile").send_keys(row["ArquivoPDF"])

    # Enviar chamado
    driver.find_element(By.ID, "btnSalvar").click()

    print("Chamado registrado:", row["Fornecedor"])
    time.sleep(5)

driver.quit()
