# -*- coding: utf-8 -*-
__author__ = 'Guilherme Cardoso de Vargas'
__version__ = 1.0
__maintainer__ = 'Guilherme Cardoso de Vargas'
__email__ = 'vargas93626@gmail.com'
__status__ = 'Development'

import pandas as pd
import glob
import os
import time
import datetime
from datetime import datetime
from selenium import webdriver 
from selenium.webdriver.firefox.options import Options
from os import listdir
from os.path import isfile, join, basename

def scrap ():
#Esconde GUI do browser
    options = Options()
    options.headless = True 

#Adiciona as preferencias
    fp = webdriver.FirefoxProfile()

    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
    fp.set_preference("browser.download.manager.showWhenStarting",False)
    fp.set_preference("browser.download.dir", "C:\\Users\\......\\WebScrapEFFICIENZA\\....")
    fp.set_preference("browser.download.folderList",2) #2 means to use the directory you specify in "browser.download.dir"

#Inicia o bot
    driver = webdriver.Firefox(options=options,executable_path="C:\\Users\\......\\WebScrapEFFICIENZA\\....geckodriver",firefox_profile=fp)

    driver.get("https://fakelink.......")
    print ("Efficienza Scrap Inicializated")


#Download arquivo
    driver.find_element_by_id("btnXls").click() 

#Tratativa para gerar timeout
    time.sleep(5)
    print ("Download concluído!!!")
    driver.close()
    driver.quit()


#Tratamento dos dados de tabelas solicitada pelo cliente
def tratmant():
    df = pd.read_html('C:\\Users\\......\\WebScrapEFFICIENZA\\....file.xls')[0]
    print("Tratando dados aguarde.")
    
    df.to_csv

    #Define (Meia noite) nas 4 primeiras linhas
    df.iloc[0:4,:1] = '00'

    df.iloc[:,2:] =  df.iloc[:,2:]/100

    df

    #Criar um dicionario das colunas
    col_mapping_dict = {c[0]:c[1] for c in enumerate(df.columns)}
    col_mapping_dict
    day1 = col_mapping_dict[2]
    day2 = col_mapping_dict[3]

    #Tratamento para nome do arquivo com a data
    day_d = pd.to_datetime(day1)
    day_d = str(day_d)
    day_d = day_d[:-9]

#Quebrar em 2 DF

#DF1
    df1 = df.iloc[:,:3]
    df1

    df1['Day'] = day1 #Definindo o dia 
    df1['Hora'] = df1['Hora'].astype(str)
    df1['Hora'] = df1['Day'] + '  ' + df1['Hora'] + ':00'  #Adicionando o dia a hora para gerar o Timestamp
    df1['Valor'] = df1.iloc[:,2:3]#Renomeia a coluna como coluna Valor

    df1 = df1.drop(['Day'],axis=1)
    df1 = df1.drop(df1.iloc[:,2:3],axis=1)

#DF2
    df2 = df.iloc[:,[0,1,3]]
    df2

    df2['Day'] = day2
    df2['Hora'] = df2['Hora'].astype(str)
    df2['Hora'] = df2['Day'] + '  ' + df2['Hora'] + ':00'
    df2['Valor'] = df2.iloc[:,2:3]

    df2 = df2.drop(['Day'],axis=1)
    df2 = df2.drop(df2.iloc[:,2:3],axis=1)

    #Converte a Hora para Timestamp (Datetime)
    df1['Hora'] = pd.to_datetime(df1['Hora'])
    df2['Hora'] = pd.to_datetime(df2['Hora'])


    # Exporta o data frame para CSV 
    df_final = pd.concat([df1, df2])
    df_final = df_final.set_index('Hora')
    df_final.to_csv('C:\\Users\\......\\WebScrapEFFICIENZA\\....\\'+"name-"+day_d+".csv")
    print("Arquivo final salvo")

path = 'C:\\Users\\......\\WebScrapEFFICIENZA\\....'

def remove():
    for f in glob.iglob(path+'/**/*.xls', recursive=True):
        os.remove(f)
        print("The file: {} is deleted!".format(f, join(basename(f))))


if __name__ == "__main__":
    print ("Loading")
    scrap()
    tratmant()
    print("Limpando arquivos antigos")
    remove()
    print("Concluído")
    print("Feche o programa.......")

