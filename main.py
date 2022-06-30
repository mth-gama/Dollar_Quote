from tkinter import *
from telnetlib import X3PAD
from time import time
from matplotlib.pyplot import text
import pyautogui as pg
import os
import openpyxl
from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
import clipboard
from matplotlib.pyplot import title
from Functions import *
import pyautogui as pg
import pandas as pd
import os


# Code created by Matheus Gama
# My gitHub: mth-gama
# Contrata EU <3

# Variables
color01 = '#025C75'
color02 = '#E7E9E9'
color03 = 'white'
site01 = 'https://www.bcb.gov.br/?bc='
site02 = 'https://www.bcb.gov.br/estabilidadefinanceira/fechamentodolar'
path_planilha = 'Cotacao.csv'
# /////////////////////////////////////////END
# Config the window
root = Tk()
root.geometry(center(root, 550, 200))
root.title('Cotação Master URL')
root.resizable(False, False)
root.config(bg=color03)
root.iconbitmap(r'img\bitmap_cotacao.ico')
# /////////////////////////////////////////END
# Dinamic Variables
data01 = StringVar()
taxa_venda01 = StringVar()
taxa_compra01 = StringVar()
data02 = StringVar()
taxa_venda02 = StringVar()
taxa_compra02 = StringVar()
data01.set('--------------')
data02.set('--------------')
taxa_compra01.set('--------------')
taxa_compra02.set('--------------')
taxa_venda01.set('--------------')
taxa_venda02.set('--------------')
# /////////////////////////////////////////END
# Functions the system


def cotar():
    drive = Chrome(
        'chromedriver.exe')
    drive.maximize_window()
    drive.get(site01)
    pg.sleep(3)
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
    if os.path.exists('Cotacao.csv'):
        pg.alert('Dados coletados com Sucesso!!')
    else:
        pg.alert(
            'Um ou mais erros impediu que o sistema coletasse as informações :(')


def abrir_CSV():
    if os.path.exists('Cotacao.csv'):
        os.startfile('Cotacao.csv')
    else:
        pg.alert('Não existe nenhum arquivo CVS criado')


def excluir_CSV():
    if os.path.exists('Cotacao.csv'):
        os.remove('Cotacao.csv')
        pg.alert('Arquivo excluido com sucesso!!')
    else:
        pg.alert('Não existe nenhum arquivo CVS criado')


def refresh():
    if os.path.exists('Cotacao.csv'):
        data_col = []
        taxa_cp = []
        taxa_vd = []
        planilha = pd.read_excel(path_planilha)
        planilha = planilha.applymap(str)

        # Function to read rows the spreatsheet
        for i, row in planilha.iterrows():
            date = planilha['Data'][i]
            date = date.rstrip()
            compra = planilha['Taxa de Compra'][i]
            compra = 'R$ '+compra
            venda = planilha['Taxa de Venda'][i]
            venda = 'R$ '+venda
            data_col.append(date)
            taxa_cp.append(compra)
            taxa_vd.append(venda)
        data01.set(data_col[0])
        data02.set(data_col[1])
        taxa_compra01.set(taxa_cp[0])
        taxa_compra02.set(taxa_cp[1])
        taxa_venda01.set(taxa_vd[0])
        taxa_venda02.set(taxa_vd[1])
    else:
        pg.alert('Não existe nenhum arquivo CVS criado')


# Containers

fr_container01 = Frame(
    root,
    width=550,
    height=100,
    bg=color01
)

fr_container02 = Frame(
    root,
    bg=color02,
    width=540,
    height=50
)
lb_mthGama = Label(
    root,
    text='Created by Matheus Gama',
    font='Verdana 8 italic'
)
fr_container01.grid(row=1, column=0)
fr_container02.grid(row=2, column=0, pady=5)
lb_mthGama.grid(row=3, column=0)
fr_container01.grid_propagate(0)
fr_container02.grid_propagate(0)
# /////////////////////////////////////////END
# Labels container 01
lb_site_title = Label(
    fr_container01,
    text='SITE',
    font='Verdana 13 bold',
    fg=color03,
    bg=color01
)
lb_site01_name = Label(
    fr_container01,
    text='Site Bacen (1)',
    font='Verdana 13',
    fg=color03,
    bg=color01,
)
lb_site02_name = Label(
    fr_container01,
    text='Site Bacen (2)',
    font='Verdana 13',
    fg=color03,
    bg=color01,
)
lb_data_title = Label(
    fr_container01,
    text='DATA',
    font='Verdana 13 bold',
    fg=color03,
    bg=color01
)
lb_data_site01 = Label(
    fr_container01,
    textvariable=data01,
    font='Verdana 13',
    fg=color03,
    bg=color01,
)
lb_data_site02 = Label(
    fr_container01,
    textvariable=data02,
    font='Verdana 13',
    fg=color03,
    bg=color01,
)
lb_taxa_compra_title = Label(
    fr_container01,
    text='TAXA COMPRA',
    font='Verdana 13 bold',
    fg=color03,
    bg=color01,

)
lb_taxa_compra_site01 = Label(
    fr_container01,
    textvariable=taxa_compra01,
    font='Verdana 13',
    fg=color03,
    bg=color01,
)
lb_taxa_compra_site02 = Label(
    fr_container01,
    textvariable=taxa_compra02,
    font='Verdana 13',
    fg=color03,
    bg=color01,
)
lb_taxa_venda_title = Label(
    fr_container01,
    text='TAXA VENDA',
    font='Verdana 13 bold',
    fg=color03,
    bg=color01,
)
lb_taxa_venda_site01 = Label(
    fr_container01,
    textvariable=taxa_venda01,
    font='Verdana 13',
    fg=color03,
    bg=color01,
)
lb_taxa_venda_site02 = Label(
    fr_container01,
    textvariable=taxa_venda02,
    font='Verdana 13',
    fg=color03,
    bg=color01,
)

lb_site_title.grid(row=0, column=0, sticky=NW, padx=1, pady=1)
lb_site01_name.grid(row=1, column=0, sticky=NW, padx=1, pady=1)
lb_site02_name.grid(row=2, column=0, sticky=NW, padx=1, pady=1)
lb_data_site01.grid(row=1, column=1, sticky=NW, padx=1, pady=1)
lb_data_site02.grid(row=2, column=1, sticky=NW, padx=1, pady=1)
lb_data_title.grid(row=0, column=1, sticky=NW, padx=1, pady=1)
lb_taxa_compra_title.grid(row=0, column=2, sticky=NW, padx=1, pady=1)
lb_taxa_compra_site01.grid(row=1, column=2, sticky=NW, padx=1, pady=1)
lb_taxa_compra_site02.grid(row=2, column=2, sticky=NW, padx=1, pady=1)
lb_taxa_venda_title.grid(row=0, column=3, sticky=NW, padx=1, pady=1)
lb_taxa_venda_site01.grid(row=1, column=3, sticky=NW, padx=1, pady=1)
lb_taxa_venda_site02.grid(row=2, column=3, sticky=NW, padx=1, pady=1)
# /////////////////////////////////////////END
# Labels Container 02
btn_refresh = Button(
    fr_container02,
    text='Refresh',
    bg='green',
    fg='white',
    command=refresh,
    height=1,
    width=12
)
btn_cotar = Button(
    fr_container02,
    text='Nova consulta',
    bg='blue',
    fg='white',
    command=cotar,
    height=1,
    width=12
)
btn_abrir_CSV = Button(
    fr_container02,
    text='Abrir arquivo',
    bg='orange',
    fg='white',
    command=abrir_CSV,
    height=1,
    width=12
)
btn_excluir_CVS = Button(
    fr_container02,
    text='Excluir CVS',
    bg='red',
    fg='white',
    command=excluir_CSV,
    height=1,
    width=12
)
btn_refresh.grid(row=0, column=0, padx=5, pady=10)
btn_cotar.grid(row=0, column=1, padx=5, pady=10)
btn_abrir_CSV.grid(row=0, column=2, padx=5, pady=10)
btn_excluir_CVS.grid(row=0, column=3, padx=5, pady=10)
root.mainloop()
