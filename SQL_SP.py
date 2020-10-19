# Importing Modules and Dependencies 
from shareplum import Site
from shareplum import Office365
from sqlalchemy import create_engine, event
import sql_connect as secret
import pyodbc
import pandas as pd
# import encrypt as E
from sys import argv

# Decoding the Sharepoint and SQL passwords
pwd = secret.sharepoint_password
pwd1 = secret.password
# Printing to the console for easy debugging.
print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()

db = 'FRP_EDW'
_SP = 'bidash'
_SPL = 'Location Mapping'

query = """  

Select 
	LocationName as [Location Name], 
	case when BusinessUnitID is null then '' else concat(BusinessUnitID, ';#', BusinessUnitName) end as [Business Unit] 
	from hris.LocationMapping

"""
def SPList(_SP, _SPL):
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
        mylist1 = site.List(_SPL)
        data = mylist1.GetListItems('All Items')
        return data
    except Exception as e : print(e)

# BusinessUnit =  pd.DataFrame(SPList(_SP, 'FRPS Business Unit'))
# OperatingUnit =  pd.DataFrame(SPList(_SP, 'FRPS Operating Unit'))
# Deal =  pd.DataFrame(SPList(_SP, 'dim Deal'))
# DataSource =  pd.DataFrame(SPList(_SP, 'dim DataSource'))

# def lookupFormat(st):
#     if '-' in st:
#         d = st.split('-')
#         # d[0] = li[]
#         s = f'{d[0]};#{d[1]}'
#         return s
#     else :
#         return ''
# print(f'Collecting data from SQL Table : {_Table}....') sql       
# connecting to SQL and reading table into a python object
engine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/{db}?driver=SQL+Server+Native+Client+11.0")
# conn = engine.connect()
# query = "SELECT COA_ID, AgencyID, DataSourceID, FRPS_DataSourceSK, TitleAccount, TitleAccountName, ChartID, FRPS_FullAccount, FRPS_StandardID  FROM frps.Chart_Agency"
df = pd.read_sql_query(query, engine)
# df = pd.read_sql_table(_Table, engine, schema=_schema)
mydata = df.to_dict(orient='records')
# lookupFields = ['BusinessUnit', 'OperatingUnit', 'DealName', 'FRPS_DataSourceSK']
# for item in mydata:
#     if 'Name' in item:
#         item['Name_'] = item.pop('Name')
#     for it in lookupFields:
#         item[it]= lookupFormat(item[it])
    
   

print(f'updating Sharepoint List... : {_SPL}....')
# Connecting to the destination sharepoint list with try 
print(mydata[0])
try:
    site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
except Exception as e : print(e)
# print(site.GetListCollection()[0])
# Writing data to sharepoint list from sql
# reading the the desitination Sharepoint list
mylist1 = site.List(_SPL)
data1 = mylist1.GetListItems('All Items')

lists = [item['Title'] for item in site.GetListCollection()]
# Retreiving all IDs of the list if it not empty
# ids = [item['ID'] for item in data1]
# Delete all items from the sharepoint list by IDs
# mylist1.UpdateListItems(ids, kind='New')
# Adding the new Data to the sharepoint list 
mylist1.UpdateListItems(data=mydata, kind='Update')
print(f'---------------Done -------------')



