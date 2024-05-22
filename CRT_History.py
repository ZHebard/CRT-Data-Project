# -*- coding: utf-8 -*-
"""
@author: Zachary Hebard
"""



import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from datetime import date
import os
import shutil


past_Dates = pd.read_excel('past_Dates.xlsx', converters={'Date':str})  

past_Dates['Date'] = past_Dates['Date'].str[:10]


past_Dates = past_Dates['Date'].tolist()


for date in past_Dates:
    yesterday = date
    try:
        service = Service(executable_path=r"C:\Users\Zachary Hebard\chromedriver")
        driver = webdriver.Chrome(service=service)

        driver.get("https://old.crt.org.mx/EstadisticasCRTweb/Informes/ExportacionesPorPais.aspx")

        time.sleep(2)

        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl03_txtValue")
        input_element.send_keys(yesterday)

        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl05_txtValue")
        input_element.send_keys (yesterday)

        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl07_ddDropDownButton")
        input_element.click()
        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl07_divDropDown_ctl00")
        input_element.click()

        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl09_ddDropDownButton")
        input_element.click()
        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl09_divDropDown_ctl00")
        input_element.click()

        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl11_ddDropDownButton")
        input_element.click()
        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl11_divDropDown_ctl00")
        input_element.click()

        input_element = driver.find_element(By.ID, "ReportViewer1_ctl04_ctl00")
        input_element.click()

        time.sleep(15)

        input_element = driver.find_element(By.ID, "ReportViewer1_ctl05_ctl04_ctl00_ButtonImg")
        input_element.click()

        time.sleep(5)

        input_element = driver.find_element(By.XPATH, '//*[@id="ReportViewer1_ctl05_ctl04_ctl00_Menu"]/div[2]/a')
        input_element.click()

        time.sleep(8)

        driver.quit()

        col_names = ["NombrePais","textbox11","Categoria","textbox14","Clase","textbox17"]
        csv_input = pd.read_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv', names = col_names)
        csv_input = csv_input.drop(
            labels=[0,1,2],
            axis = 0)
        csv_input = csv_input.rename(columns={'NombrePais': 'Country', 'textbox11': 'Total Liters', 'Categoria': 'Category', 'textbox14': 'Category Liters', 'Clase': 'Class', 'textbox17': 'Class Liters'})
        
        t3 = csv_input.pivot(values='Class Liters', index='Country', columns=['Category', 'Class']).swaplevel(0,1,axis=1)
        t3.columns = t3.columns.map('_'.join)


        t2=pd.pivot_table(csv_input, values='Category Liters', index='Country', columns='Category', aggfunc='max')


        t1=pd.pivot_table(csv_input, values='Total Liters', index='Country', aggfunc='max')



        t4 = pd.merge(t1, t2, on='Country', how='outer')
        t5 = pd.merge(t3, t4, on='Country', how='outer')
        t5['Date']=yesterday
        t5.to_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv')


        column_names = [
            
            ["BLANCO_TEQUILA"],
            ["EXTRA AÑEJO_TEQUILA 100% DE AGAVE"],
            ["EXTRA AÑEJO_TEQUILA"],
            ["JOVEN_TEQUILA 100% DE AGAVE"],
            ["JOVEN_TEQUILA"],
            ["AÑEJO_TEQUILA 100% DE AGAVE"],
            ["BLANCO_TEQUILA 100% DE AGAVE"],
            ["REPOSADO_TEQUILA 100% DE AGAVE"],
            ["AÑEJO_TEQUILA"],
            ["REPOSADO_TEQUILA"],
            
]

        flat_column_names = [item for sublist in column_names for item in sublist]

        missing_columns = set(flat_column_names) - set(t5.columns)

        # Update only the missing columns in df
        t5 = t5.assign(**{col: pd.Series(dtype=object) for col in missing_columns})

        t5.to_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv')


        shutil.move(r"C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv", r"C:\Users\Zachary Hebard\CRT_Data\ReporteCategoriaClasePais.csv")

        os.chdir(r"C:\Users\Zachary Hebard\CRT_Data")
        for file in os.listdir():
            if file == 'ReporteCategoriaClasePais.csv':
                continue
   
        
        source= (r"ReporteCategoriaClasePais.csv")
        fileName = yesterday
        #.strftime('%d-%m-%Y')
        fileName = fileName+".csv"
        dest = os.path.join(os.path.dirname(source), fileName)

        os.rename(source, dest)
    except Exception as e:
        print(f"Error scraping data for {yesterday}: {e}")
        os.chdir(r"C:\Users\Zachary Hebard\Downloads")
        os.remove(r"ReporteCategoriaClasePais.csv")
   
    pass  
    








    






