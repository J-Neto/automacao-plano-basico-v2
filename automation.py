from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains
import time
import pandas as pd
import math


def openBrowser():

    # Using ChromeOptions() to supress the "bluetooth_adapter_winrt.cc" error
    # We aren't using bluetooth, so its fine to supress this error message
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # Looking for chromedriver
    driver = webdriver.Chrome(options=options, service=Service(
        ChromeDriverManager().install()))

    # Opening Chrome Browser
    url = "http://sistemas.anatel.gov.br/se/public/view/b/srd.php"
    driver.get(url)
    time.sleep(5)

    return driver


def registersPerPage(driver, n):

    # Getting the number of registers field
    qtd_field = driver.find_element(By.XPATH, '//*[@id="rpp"]')

    # Perfoming a double click
    action = ActionChains(driver)
    action.double_click(qtd_field).perform()

    # Defining the number of registers and pressing ENTER (RETURN button)
    qtd_field.send_keys('250' + Keys.RETURN)

    time.sleep(3)


def filterService(driver, service_name):

    # Clicking on "Filtrar" element
    driver.find_element(By.XPATH, '//*[@id="tblFilter"]/span[5]').click()
    time.sleep(10)

    # Type the "Serviço" search input, searching for the service name and pressing ENTER (RETURN button)
    driver.find_element(
        By.XPATH, '//*[@id="fc_6"]').send_keys(service_name + Keys.RETURN)
    time.sleep(10)


def nextPage(driver):
    element = driver.find_element(By.XPATH, '//*[@id="nextPageOffset"]')
    driver.execute_script("arguments[0].scrollIntoView()", element)
    driver.execute_script("arguments[0].click()", element)
    time.sleep(10)


def getDataToTable(driver, steps):
    for i in range(steps):
        # Getting Table
        tbl = driver.find_element(
            By.XPATH, '//*[@id="aplTbl"]').get_attribute('outerHTML')

        # Converting to list of pandas dataframe
        df_raw = pd.read_html(tbl)

        # If is the first time, we just create the dataframe
        if (i == 0):
            # Getting the first element of list, which is our data
            df = df_raw[0]

        # If is not:
        #   we get the existent dataframe,
        #   then join itself with the new dataframe that contains the new data of the new page
        else:
            # Getting the first element of list, which is our data
            df_raw = df_raw[0]

            # Then joining with the existent dataframe 'df'
            frames = [df, df_raw]
            df = pd.concat(frames)

        if (i != steps - 1):
            time.sleep(10)
            nextPage(driver)

    return df


def getSteps(driver):
    # Getting the span that contains the total number of registers
    raw = driver.find_element(By.XPATH, '//*[@id="tblFilter"]/span[1]').text

    # The 'raw' will be on this format: 'N total de registros'
    # Spliting into a list with 'total' as the separator.
    str_qtd = raw.split('total')

    # Now our list is: ['N ', 'total de registros']
    # We get the first element, then strip removing the white spaces
    str_qtd = str_qtd[0].strip()

    # Now we have 'N'. So let's convert into a int. This is our number of total registers
    qtd = int(str_qtd)

    # The page show just 250 itens per time, so let's calculate how many times the loop will be running
    n = qtd/250
    steps = math.ceil(n)

    return steps


def removeColumnsDf(df):
    # Droping columns
    df.drop('Finalidade', inplace=True, axis=1)
    df.drop('Num Serviço', inplace=True, axis=1)
    df.drop('Local Especifico', inplace=True, axis=1)
    df.drop('Categoria da Estação', inplace=True, axis=1)
    df.drop('Fase', inplace=True, axis=1)
    df.drop('Data', inplace=True, axis=1)
    df.drop('ERP', inplace=True, axis=1)
    df.drop('HCI', inplace=True, axis=1)
    df.drop('ID Estação Principal', inplace=True, axis=1)

    return df


def tableTreatment(df):

    # Removing custom columns
    df = removeColumnsDf(df)

    # Replacing NaN values with blank space
    df.fillna("", inplace=True)

    # On column "Entidade", replace "" with "CANAL VAGO"
    df["Entidade"].replace(r'^\s*$', "CANAL VAGO", regex=True, inplace=True)

    # Getting index of blank rows of 'Ações' column
    index0Row = df[df['Ações'] == ""].index

    # Droping blank rows
    df.drop(index0Row, inplace=True)

    # Reseting df index
    df.reset_index(drop=True, inplace=True)

    return df


def copyPaste(driver):

    steps = getSteps(driver)

    # For loop
    df = getDataToTable(driver, 2)

    # Table treatment
    df = tableTreatment(df)

    # Saving data to a Table, removing index
    df.to_excel(
        "C:/Mirante/Projects/automacao-plano-basico-v2/plano-basico.xlsx", index=False)


def search(driver):

    # Define 250 as the number os registers per page
    registersPerPage(driver, 250)

    # Filtering the service name as "TV"
    filterService(driver, "TV")


def closeBrowser(driver):

    # Wait 5 seconds than close the Browser
    time.sleep(5)
    driver.close()


chrome = openBrowser()
search(chrome)
df = copyPaste(chrome)
closeBrowser(chrome)
