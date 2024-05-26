# -*- coding: utf-8 -*-
"""
Created on Sat May 25 21:25:19 2024

@author: Zachary Hebard
"""

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

#setting the web adress for selenium to navigate 
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
csv_input = csv_input.rename(columns={'NombrePais': 'Country_Spanish', 'textbox11': 'Total Liters', 'Categoria': 'Category', 'textbox14': 'Category Liters', 'Clase': 'Class', 'textbox17': 'Class Liters'})


#selecting page elements and entering date information or selecting check boxes to get all 
#wanted data, downloaded data as .csv
t3 = csv_input.pivot(values='Class Liters', index='Country_Spanish', columns=['Category', 'Class']).swaplevel(0,1,axis=1)
t3.columns = t3.columns.map('_'.join)
t2=pd.pivot_table(csv_input, values='Category Liters', index='Country_Spanish', columns='Category', aggfunc='max')
t1=pd.pivot_table(csv_input, values='Total Liters', index='Country_Spanish', aggfunc='max')

#Creating DFs T4 and T5 to merge T1-T3 to dinish the conversion from long to wide data and adding a new column
#called date to represent the data of which the given data represents. Re-write T5 back to same .csv file
t4 = pd.merge(t1, t2, on='Country_Spanish', how='outer')
t5 = pd.merge(t3, t4, on='Country_Spanish', how='outer')
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

# Update only the missing columns in df
t5 = t5.assign(**{col: pd.Series(dtype=object) for col in missing_columns})
t5.to_csv(r'C:\Users\Zachary Hebard\Downloads\ReporteCategoriaClasePais.csv')

#Creating a dictionary linking spanish and english country spellings
country_map = {
    'ALBANIA':'ALBANIA',
    'ARGELIA':'ALGERIA',
    'ANDORRA':'ANDORRA',
    'ANGOLA':'ANGOLA',
    'ARGENTINA':'ARGENTINA',
    'ARMENIA ':'ARMENIA',
    'ARUBA':'ARUBA',
    'AUSTRALIA':'AUSTRALIA',
    'AUSTRIA':'AUSTRIA',
    'AZERBAIJAN ':'AZERBAIJAN',
    'BAHAMAS':'BAHAMAS',
    'BAHREIN':'BAHRAIN',
    'BANGLADESH':'BANGLADESH',
    'BARBADOS':'BARBADOS',
    'BIELORRUSIA':'BELARUS',
    'BELGICA':'BELGIUM',
    'BELICE':'BELIZE',
    'BENIN':'BENIN',
    'BERMUDAS':'BERMUDA',
    'BOLIVIA':'BOLIVIA',
    'BOSNIA':'BOSNIA',
    'BOTSWANA':'BOTSWANA',
    'BRASIL':'BRAZIL',
    'TERRITORIOS BRITANICOS DEL OCEANO INDICO':'BRITISH INDIAN OCEAN TERRITORIES',
    'ISLAS VIRGENES BRITANICAS':'BRITISH VIRGIN ISLANDS',
    'BULGARIA':'BULGARIA',
    'CAMBOYA':'CAMBODIA',
    'CAMERUN':'CAMEROON',
    'CANADA':'CANADA',
    'ISLAS CAIMAN  ':'CAYMAN ISLANDS',
    'CHILE':'CHILI',
    'CHINA':'CHINA',
    'COLOMBIA':'COLOMBIA',
    'COSTA RICA':'COSTA RICA',
    'CROACIA':'CROATIA',
    'CUBA':'CUBA',
    'CURAZAO':'CURACAO',
    'CHIPRE':'CYPRUS',
    'REPUBLICA CHECA':'CZECH REPUBLIC',
    'REPUBLICA DEMOCRATICA DEL CONGO':'DEMOCRATIC REPUBLIC OF CONGO',
    'DINAMARCA':'DENMARK',
    'DJIBOUTI':'DJIBOUTI',
    'REPUBLICA DOMINICANA':'DOMINICAN REPUBLIC',
    'ECUADOR':'ECUADOR',
    'ANGUILA':'EEL',
    'EGIPTO':'EGYPT',
    'ESTONIA':'ESTONIA',
    'ETIOPIA':'ETHIOPIA',
    'ESTADO FEDERADO DE MICRONESIA':'FEDERATED STATE OF MICRONESIA',
    'FIDJI':'FIDJI',
    'FINLANDIA':'FINLAND',
    'FRANCIA':'FRANCE',
    'POLINESIA FRANCESA':'FRENCH POLYNESIA',
    'GEORGIA':'GEORGIA',
    'ALEMANIA':'GERMANY',
    'GHANA':'GHANA',
    'GIBRALTAR':'GIBRALTAR',
    'GRECIA':'GREECE',
    'GUADALUPE':'GUADELOUPE',
    'GUAM EUA':'GUAM USA',
    'GUATEMALA':'GUATEMALA',
    'GUYANA':'GUYANA',
    'HONDURAS':'HONDURAS',
    'HONG KONG':'HONG KONG',
    'HUNGRIA':'HUNGARY',
    'ISLANDIA':'ICELAND',
    'INDIA':'INDIA',
    'INDONESIA':'INDONESIA',
    'IRAK':'IRAQ',
    'IRLANDA':'IRELAND',
    'ISRAEL':'ISRAEL',
    'ITALIA':'ITALY',
    'JAMAICA':'JAMAICA',
    'JAPON':'JAPAN',
    'JORDANIA':'JORDAN',
    'KAZAKHSTAN':'KAZAKHSTAN',
    'KENYA':'KENYA',
    'LETONIA':'LATVIA',
    'LIBANO':'LEBANON',
    'LITUANIA':'LITHUANIA',
    'LUXEMBURGO':'LUXEMBOURG',
    'MACAO':'MACAU',
    'MALASIA':'MALAYSIA',
    'MALDIVAS':'MALDIVES',
    'MALTA':'MALT',
    'MARTINICA':'MARTINIQUE',
    'MAURICIO':'MAURICIO',
    'MEXICO':'MEXICO',
    'MOLDAVIA':'MOLDOVA',
    'MONACO':'MONACO',
    'MARRUECOS':'MOROCCO',
    'MOZAMBIQUE':'MOZAMBIQUE',
    'PAISES BAJOS':'NETHERLANDS',
    'ANTILLAS NEERLANDESAS':'NETHERLANDS ANTILLES',
    'NUEVA CALEDONIA':'NEW CALEDONIA',
    'NUEVA ZELANDIA':'NEW ZEALAND',
    'NICARAGUA':'NICARAGUA',
    'NIGERIA':'NIGERIA',
    'ISLAS VIRGENES NORTEAMERICANAS':'NORTH AMERICAN VIRGIN ISLANDS',
    'COREA DEL NORTE':'NORTH KOREA',
    'NORUEGA':'NORWAY',
    'PANAMA':'PANAMA',
    'PARAGUAY':'PARAGUAY',
    'PERU':'PERU',
    'FILIPINAS':'PHILIPPINES',
    'POLONIA':'POLAND',
    'PORTUGAL':'PORTUGAL',
    'PUERTO RICO':'PUERTO RICO',
    'QATAR':'QATAR',
    'RUMANIA':'ROMANIA',
    'RUSIA':'RUSSIA',
    'REPUBLICA RUANDESA':'RWANDASE REPUBLIC',
    'SAN CRISTOBAL Y NIEVES':'SAINT KITTS AND NEVIS',
    'SAN MARTÍN':'SAN MARTIN',
    'ARABIA SAUDITA':'SAUDI ARABIA',
    'SERBIA':'SERBIA',
    'SEYCHELLES':'SEYCHELLES',
    'SIERRA LEONA':'SIERRA LEONE',
    'SINGAPUR':'SINGAPORE',
    'REPUBLICA ESLOVACA':'SLOVAK REPUBLIC',
    'ESLOVENIA':'SLOVENIA',
    'SUDAFRICA':'SOUTH AFRICA',
    'COREA DEL SUR':'SOUTH KOREA',
    'ESPAÑA':'SPAIN',
    'SRI LANKA':'SRI LANKA',
    'SANTA LUCIA':'ST. LUCIA',
    'SAN VICENTE Y LAS GRANADINAS':'ST. VINCENT AND THE GRENADINES',
    'SURINAME':'SURINAME',
    'SUECIA':'SWEDEN',
    'SUIZA':'SWISS',
    'SIRIA':'SYRIA',
    'TAIWAN':'TAIWAN',
    'TAILANDIA':'THAILAND',
    'EL SALVADOR':'THE SAVIOR',
    'TRINIDAD Y TOBAGO':'TRINIDAD AND TOBAGO',
    'TURQUIA':'TURKEY',
    'TURCAS Y CAICOS':'TURKS AND CAICOS',
    'UCRANIA':'UKRAINE',
    'EMIRATOS ARABES UNIDOS':'UNITED ARAB EMIRATES',
    'REINO UNIDO DE LA GRAN BRETAÑA E IRLANDA DEL NORTE':'United Kingdom',
    'URUGUAY':'URUGUAY',
    'ESTADOS UNIDOS DE AMERICA':'USA',
    'UZBEJISTAN':'UZBEJISTAN',
    'VENEZUELA':'VENEZUELA',
    'VIETNAM':'VIETNAM'
}

#Reset index, add column of Enlhish county spelling using country_map dictionary and reconfigure df
t5 = t5.reset_index()
t5['Country'] = t5['Country_Spanish'].apply(lambda x: country_map.get(x, None))
t5 = t5.loc[:, ['Country', 'AÑEJO_TEQUILA 100% DE AGAVE', 'BLANCO_TEQUILA 100% DE AGAVE', 'REPOSADO_TEQUILA 100% DE AGAVE', 
                'EXTRA AÑEJO_TEQUILA 100% DE AGAVE',	'BLANCO_TEQUILA',	'REPOSADO_TEQUILA',	'JOVEN_TEQUILA',	
                'JOVEN_TEQUILA 100% DE AGAVE',	'Total Liters',	'TEQUILA',	'TEQUILA 100% DE AGAVE', 'Date']]  
t5 = t5.set_index('Country')
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
connection_string = f"postgresql://postgres:el_pandillo@localhost:5433/CRT_Data"
engine = create_engine(connection_string)
table_name = 'exports'
t5.to_sql(table_name, engine, if_exists='append', index=True)
