# -*- coding: utf-8 -*-
"""
@author: Zachary Hebard
"""


#importing libaries
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from datetime import date
import os
import shutil

#Read in excel sheet with historical dates needed for data pull
#Covert Date field to a string and extracting first 9 characters
#converting past_Dates to a list
past_Dates = pd.read_excel('past_Dates.xlsx', converters={'Date':str})  
past_Dates['Date'] = past_Dates['Date'].str[:10]
past_Dates = past_Dates['Date'].tolist()

#start for loop to cycle through all dates in past_Dates
#assigning the first date to the variable yesterday
for date in past_Dates:
    yesterday = date
    #Setting up chrome driver with selenium to automate chrome
    try:
        service = Service(executable_path=r"C:\Users\Zachary Hebard\chromedriver")
        driver = webdriver.Chrome(service=service)
        #setting site link for slenium to navigate
        driver.get("https://old.crt.org.mx/EstadisticasCRTweb/Informes/ExportacionesPorPais.aspx")
        #selecting page elements and entering date information or selecting check boxes to get all 
        #wanted data, downloaded data as .csv
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
        
        #Reading in above downloaded .csv file as csv_input and droping unwanted data and renaming columns
        col_names = ["NombrePais","textbox11","Categoria","textbox14","Clase","textbox17"]
        csv_input = pd.read_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv', names = col_names)
        csv_input = csv_input.drop(
            labels=[0,1,2],
            axis = 0)
        csv_input = csv_input.rename(columns={'NombrePais': 'Country', 'textbox11': 'Total Liters', 'Categoria': 'Category', 'textbox14': 'Category Liters', 'Clase': 'Class', 'textbox17': 'Class Liters'})
        
        #creating new DFs as T1-T3 to create pivot to change the df format fron long to wide.
        t3 = csv_input.pivot(values='Class Liters', index='Country', columns=['Category', 'Class']).swaplevel(0,1,axis=1)
        t3.columns = t3.columns.map('_'.join)
        t2=pd.pivot_table(csv_input, values='Category Liters', index='Country', columns='Category', aggfunc='max')
        t1=pd.pivot_table(csv_input, values='Total Liters', index='Country', aggfunc='max')

        #Creating DFs T4 and T5 to merge T1-T3 to dinish the conversion from long to wide data and adding a new column
        #called date to represent the data of which the given data represents. Re-write T5 back to same .csv file
        t4 = pd.merge(t1, t2, on='Country', how='outer')
        t5 = pd.merge(t3, t4, on='Country', how='outer')
        t5['Date']=yesterday
        t5.to_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv')

        #FOR SQL IMPORT- Create a list of all needed column names even if not reprecented in the dates data. 
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
        #Finding .csv files missing columns from column_names list and adding them to the DF
        #overwrite T5 and re-wrtire back to .csv file
        flat_column_names = [item for sublist in column_names for item in sublist]
        missing_columns = set(flat_column_names) - set(t5.columns)
        t5 = t5.assign(**{col: pd.Series(dtype=object) for col in missing_columns})
        t5.to_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv')

        #move.csv file from downaloads into a folder called CRT_Data to hold all .csv files
        shutil.move(r"C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv", r"C:\Users\Zachary Hebard\CRT_Data\ReporteCategoriaClasePais.csv")

        #Change working dirctory and find default named .csv file and rename with the date representing the data
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
        
    #Creating and exception to aid in automation to skip dates where no exports were made resulting in a blnak .csv that would error out.
    #Changing directory back to downaload, finding the default named .csv and deleting it
    except Exception as e:
        print(f"Error scraping data for {yesterday}: {e}")
        os.chdir(r"C:\Users\Zachary Hebard\Downloads")
        os.remove(r"ReporteCategoriaClasePais.csv")
   
    pass  
    








    






