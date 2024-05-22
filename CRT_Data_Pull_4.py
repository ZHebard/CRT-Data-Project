
# -*- coding: utf-8 -*-
"""
@author: Zachary Hebard
"""


#importing libaries
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from datetime import date
from datetime import timedelta
import pandas as pd
import os
import shutil
from sqlalchemy import create_engine 

#Bring in todays date and subtract one day to get yesterdays date saved ar raw_yesterday
#Format date correctly for CRT input and save as yesterday. Format date as
#US format for filename and DATE columns.
raw_Today = date.today()
raw_Yesterday = raw_Today - timedelta(days = 1)
yesterday = raw_Yesterday.strftime('%d/%m/%Y')
fileName = raw_Yesterday.strftime('%Y-%m-%d')

#Setting up chrome driver with selenium to automate chrome    
service = Service(executable_path=r"C:\Users\Zachary Hebard\chromedriver")
driver = webdriver.Chrome(service=service)
driver.get("https://old.crt.org.mx/EstadisticasCRTweb/Informes/ExportacionesPorPais.aspx")
time.sleep(2)

#selecting page elements and entering date information or selecting check boxes to get all 
#wanted data, downloaded data as .csv
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

#selecting page elements and entering date information or selecting check boxes to get all 
#wanted data, downloaded data as .csv
col_names = ["NombrePais","textbox11","Categoria","textbox14","Clase","textbox17"]
csv_input = pd.read_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv', names = col_names)
csv_input = csv_input.drop(
    labels=[0,1,2],
    axis = 0)
csv_input = csv_input.rename(columns={'NombrePais': 'Country', 'textbox11': 'Total Liters', 'Categoria': 'Category', 'textbox14': 'Category Liters', 'Clase': 'Class', 'textbox17': 'Class Liters'})

#selecting page elements and entering date information or selecting check boxes to get all 
#wanted data, downloaded data as .csv
t3 = csv_input.pivot(values='Class Liters', index='Country', columns=['Category', 'Class']).swaplevel(0,1,axis=1)
t3.columns = t3.columns.map('_'.join)
t2=pd.pivot_table(csv_input, values='Category Liters', index='Country', columns='Category', aggfunc='max')
t1=pd.pivot_table(csv_input, values='Total Liters', index='Country', aggfunc='max')

#Creating DFs T4 and T5 to merge T1-T3 to dinish the conversion from long to wide data and adding a new column
#called date to represent the data of which the given data represents. Re-write T5 back to same .csv file
t4 = pd.merge(t1, t2, on='Country', how='outer')
t5 = pd.merge(t3, t4, on='Country', how='outer')
t5['Date']=fileName

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
source = (r"ReporteCategoriaClasePais.csv")
fileName = fileName+".csv"
dest = os.path.join(os.path.dirname(source), fileName)
os.rename(source, dest)

#Converting all data types to strings for import into SQL server
t5['BLANCO_TEQUILA'] = t5['BLANCO_TEQUILA'].astype(str)
t5['BLANCO_TEQUILA 100% DE AGAVE'] = t5['BLANCO_TEQUILA 100% DE AGAVE'].astype(str)
t5['REPOSADO_TEQUILA 100% DE AGAVE'] = t5['REPOSADO_TEQUILA 100% DE AGAVE'].astype(str)
t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'] = t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'].astype(str)
t5['JOVEN_TEQUILA'] = t5['JOVEN_TEQUILA'].astype(str)
t5['REPOSADO_TEQUILA'] = t5['REPOSADO_TEQUILA'].astype(str)
t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'] = t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'].astype(str)
t5['JOVEN_TEQUILA 100% DE AGAVE'] = t5['JOVEN_TEQUILA 100% DE AGAVE'].astype(str)
t5['TEQUILA'] = t5['TEQUILA'].astype(str)
t5['TEQUILA 100% DE AGAVE'] = t5['TEQUILA 100% DE AGAVE'].astype(str)
t5['AÑEJO_TEQUILA'] = t5['AÑEJO_TEQUILA'].astype(str)
t5['EXTRA AÑEJO_TEQUILA'] = t5['EXTRA AÑEJO_TEQUILA'].astype(str)
t5['AÑEJO_TEQUILA 100% DE AGAVE'] = t5['AÑEJO_TEQUILA 100% DE AGAVE'].astype(str)
t5['BLANCO_TEQUILA'] = t5['BLANCO_TEQUILA'].str.replace(',', '')
t5['BLANCO_TEQUILA 100% DE AGAVE'] = t5['BLANCO_TEQUILA 100% DE AGAVE'].str.replace(',', '')
t5['REPOSADO_TEQUILA 100% DE AGAVE'] = t5['REPOSADO_TEQUILA 100% DE AGAVE'].str.replace(',', '')
t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'] = t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'].str.replace(',', '')
t5['JOVEN_TEQUILA'] = t5['JOVEN_TEQUILA'].str.replace(',', '')
t5['REPOSADO_TEQUILA'] = t5['REPOSADO_TEQUILA'].str.replace(',', '')
t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'] = t5['EXTRA AÑEJO_TEQUILA 100% DE AGAVE'].str.replace(',', '')
t5['JOVEN_TEQUILA 100% DE AGAVE'] = t5['JOVEN_TEQUILA 100% DE AGAVE'].str.replace(',', '')
t5['Total Liters'] = t5['Total Liters'].str.replace(',', '')
t5['TEQUILA'] = t5['TEQUILA'].str.replace(',', '')
t5['TEQUILA 100% DE AGAVE'] = t5['TEQUILA 100% DE AGAVE'].str.replace(',', '')
t5['AÑEJO_TEQUILA'] = t5['AÑEJO_TEQUILA'].str.replace(',', '')
t5['EXTRA AÑEJO_TEQUILA'] = t5['EXTRA AÑEJO_TEQUILA'].str.replace(',', '')
t5['AÑEJO_TEQUILA 100% DE AGAVE'] = t5['AÑEJO_TEQUILA 100% DE AGAVE'].str.replace(',', '')

#establishing connection with SQL server
connection_string = f"postgresql://user:pass@localhost:port/CRT_Data"
engine = create_engine(connection_string)

#Identify table for import and import fianl T5 df.
table_name = 'exports'
t5.to_sql(table_name, engine, if_exists='append', index=True)
    




