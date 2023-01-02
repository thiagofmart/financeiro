from selenium.webdriver import Firefox
from selenium.webdriver.common.by import By
import pandas as pd
from time import sleep


usuario = ''
senha = ''

browser = Firefox()
browser.implicitly_wait(5)

def logar():
    browser.get("https://fs4.iclass.com.br/iclassfs/login/login.seam")
    browser.find_element(By.ID, "login:username").send_keys(usuario)
    browser.find_element(By.ID, "login:password").send_keys(senha)
    browser.find_element(By.ID, "login:id_login").click()

def get_data():
    data = pd.read_excel("data.xlsx")
    return data


def criar_OS(cliente):
    try:
        browser.get("https://fs4.iclass.com.br/iclassfs/restrict/ordemservico_criacao_edit.seam")
        # CLIENTE
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        browser.find_element(By.ID, "formEdit:nomeTitularDecoration:hidelink").click()
        browser.find_element(By.ID, "formSearchParamslovContrato:paramLovOptionText1").send_keys(cliente)
        browser.find_element(By.ID, "formSearchParamslovContrato:find").click()
        browser.find_element(By.ID, "formSearchResultslovContrato:resultTable:0:selecionar").click()
        browser.find_element(By.ID, "formControllovAmbiente:imgFechar").click() # Close
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        #ATENDIMENTO INICIAL
        browser.find_element(By.ID, "formEdit:tipoOsDecoration:hidelink").click()
        browser.find_element(By.ID, "formSearchParamslovTipoOs:paramLovOptionText1").send_keys("Atendimento inicial")
        browser.find_element(By.ID, "formSearchParamslovTipoOs:find").click()
        browser.find_element(By.ID, "formSearchResultslovTipoOs:resultTable:0:selecionar").click()
        # DATA DE AGENDAMENTO
        # CRIAR
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        sleep(1)
        browser.find_element(By.ID, "formEdit:criar").click()
        sleep(3)
    except:
        criar_OS(cliente)

logar()

def iniciar():
    df = get_data()
    for row in df.iterrows():
        if row[1]['EQUIPE']=='N':
            for i in range(row[1]['QTDE OS']):
                criar_OS(row[1]['CLIENTE'])
