from telnetlib import X3PAD
from time import time
from matplotlib.pyplot import text
import pyautogui as pg
import os
import openpyxl
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
import clipboard

# Code created by Matheus Gama
# My gitHub: mth-gama
# Contrata EU <3


# Variables
site01 = 'https://www.bcb.gov.br/?bc='
site02 = 'https://www.bcb.gov.br/estabilidadefinanceira/fechamentodolar'
# ///////////////////////////////////////////////////////////////////////////END

# Functions


def acesso_site():
    drive = Chrome(
        'chromedriver.exe')
    drive.maximize_window()
    drive.get(site01)

    elemento1 = drive.find_element(
        By.XPATH, '//*[@id="home"]/div/div[1]/div[1]/div/cotacao/table[1]/tbody/tr[1]/td[1]/span')
    data01 = elemento1.text

    elemento2 = drive.find_element(
        By.XPATH, '//*[@id="home"]/div/div[1]/div[1]/div/cotacao/table[1]/tbody/tr[1]/td[2]/span')
    taxa_venda01 = elemento2.text

    elemento3 = drive.find_element(
        By.XPATH, '//*[@id="home"]/div/div[1]/div[1]/div/cotacao/table[1]/tbody/tr[1]/td[3]/span')
    taxa_compra01 = elemento3.text
    drive.get(site02)

    pg.sleep(5)
    pg.moveTo(428, 483)
    pg.click()
    pg.click()
    pg.click()
    pg.hotkey('ctrl', 'c')
    data02 = clipboard.paste()

    pg.doubleClick(543, 483)
    pg.hotkey('ctrl', 'c')
    taxa_compra02 = clipboard.paste()

    pg.doubleClick(635, 484)
    pg.hotkey('ctrl', 'c')
    taxa_venda02 = clipboard.paste()
    #  If file exist the sintax remove and create new file
    if os.path.exists(r'Cotacao.csv'):
        os.remove(r'Cotacao.csv')
        book = openpyxl.Workbook()
        cotacao_page = book['Sheet']
        cotacao_page.append(
            ['Site', 'Data', 'Taxa de Compra', 'Taxa de Venda'])
        cotacao_page.append(
            ['Site Bacen (1)', data01, taxa_compra01, taxa_venda01])
        cotacao_page.append(
            ['Site Bacen (2)', data02, taxa_compra02, taxa_venda02])
        book.save('Cotacao.csv')
    else:
        book = openpyxl.Workbook()
        cotacao_page = book['Sheet']
        cotacao_page.append(
            ['Site', 'Data', 'Taxa de Compra', 'Taxa de Venda'])
        cotacao_page.append(
            ['Site Bacen (1)', data01, taxa_compra01, taxa_venda01])
        cotacao_page.append(
            ['Site Bacen (2)', data02, taxa_compra02, taxa_venda02])
        book.save('Cotacao.csv')


# ////////////////////////////////////////////////////////////////////////////END
acesso_site()
