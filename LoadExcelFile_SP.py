# Importing Modules and Dependencies 
from shareplum import Site
from shareplum import Office365
from sqlalchemy import create_engine, event
import sql_connect as secret
import pyodbc
import pandas as pd
import numpy as np
# import encrypt as E
from sys import argv

# Decoding the Sharepoint and SQL passwords
pwd = secret.sharepoint_password
pwd1 = secret.password
# Printing to the console for easy debugging.
print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()

def SPList(_SP, _SPL):
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
        mylist1 = site.List(_SPL)
        data = mylist1.GetListItems('All Items')
        return data
    except Exception as e : print(e)

df = pd.read_excel('../../../Downloads/SharePoint_List_Glossary.xlsx').replace(np.nan, '', regex=True)
mydata = df.to_dict('records') 

    
    
_SP = 'bidash'
_SPL = "Sharepoint Lists"

print(f'updating Sharepoint List... : {_SPL}....')
# Connecting to the destination sharepoint list with try 
print(mydata[0])
# try:
site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
# except Exception as e : print(e)
mylist1 = site.List(_SPL)
data1 = mylist1.GetListItems('All Items')

# lists = [item['Title'] for item in site.GetListCollection()]

# Adding the new Data to the sharepoint list 
mylist1.UpdateListItems(data=mydata, kind='New')
print(f'---------------Done -------------')






