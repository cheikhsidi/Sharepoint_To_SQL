# Importing Dependencies and modules
from shareplum import Site
from shareplum import Office365
from sqlalchemy import create_engine, event
import pyodbc
import sql_connect as secret
import pandas as pd
import numpy as np
import time
from timeit import default_timer as timer

start = timer()

# setting up Passwords 
sh_pwd = secret.sharepoint_password
sq_pwd = secret.password

# setting up the authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=sh_pwd).GetCookies()

# Setting up the connection to sharepoint site
site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/bidash/', authcookie=authcookie)
# Connecting to the SharePoint_Admin list
sp_list1 = site1.List('Sharepoint_Admin')
data1 = sp_list1.GetListItems('All Items')

# ------------------------------------------------------------------------------------
def PushToSql(df, db, Insert_Methode):
  '''This function pushes the supplaied daatfarme to the supplied db and insert methode (replace or append)'''
  engine = create_engine(f"mssql+pyodbc://{secret.user}:{sq_pwd}@{secret.server}:1433/{db}?driver=SQL+Server+Native+Client+11.0", fast_executemany=True)
  print(f'inserting {li} into {db}')
  df.to_sql(dest, engine, schema=schema, if_exists = Insert_Methode, index=False)
    
def largeList(sp_list):
    header = sp_list.GetListItems('All Items', rowlimit=1)
    # getting the list of fields to be inserted to SQL
    # retreiving the first ID
    i = int(header[0]['ID'])
    n = i + 5000
    m = 5000
    count = 0
    # looping over the 5000 chuncks at time because of the limit of SharePoint
    while m == 5000:
        print(i, n, m, count)
        query = {'Where': ['And', ('Geq', 'ID', str(i)), ('Lt', 'ID', str(n))]}
        print(query)
        dt = sp_list.GetListItems(viewname='All Items', query=query) 
        df = pd.DataFrame(dt)
        if count == 0:
            # data = dt
            data = df.replace(np.nan, '', regex=True).to_dict('records')
        else :
            print(f"{count+1} batch...{len(df)}")
            # data.extend(dt)
            data.extend(df.replace(np.nan, '', regex=True).to_dict('records'))

        i = i + 5000
        n = n + 5000
        m = len(dt)
        count += 1
    print(f"largeList count : {len(data)}")
    return data
   
def smallListPush(Insert_Methode, data):
  '''This function pushes small sharepoint list less than 5000 items, require the methode (replace, append) replace to replace 
  drop sql table and recreate it, append to add to an existing table.
  another argument is required the data object that need to be pushed to the table '''
  # Creating a dataframe of the list
  df = pd.DataFrame(data)
  # writing the dataframe to sql
  if len(db.split(',')) > 1:
    for d in db.split(','):
      d = d.strip()
      #   Inserting Data into all datatabses
      PushToSql(df, d, Insert_Methode)
  else :
    #   Inserting Data into all datatabses
    PushToSql(df, db, Insert_Methode)
# ------------------------------------------------------------------------------------
# Reading the source and destination tables
for item in data1:
  li = item['SharePoint_List']
  db = item['Destination_DB']
  siteName = item['Site_Name']
  schema = item['_Schema']
  dest = item['SQL_Table_Name']

  try:
    # establishing connection to the sharepoint site where the target list lives.
    site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{siteName}/', authcookie=authcookie)
    # Reading the Sahrepoint lists in a python object and inserting it to SQl
    sp_list = site.List(li)
    # retreiving the header of the list
    data = sp_list.GetListItems('All Items')
    if type(data) is list:
      smallListPush('replace', data)
    else :
      # retreiving the header of the list
      print("large list skipping the try...pushing...")
      data = largeList(sp_list)
      smallListPush('replace', data) 
  #   # If the connection fails print the Error to the console 
  except Exception as e: print(e)
end = timer()
tm = end - start
print (f"Time elapsed to insert items is : {round(tm/60, 2)} minutes")   
print(f"----------------- Done --------------------")


