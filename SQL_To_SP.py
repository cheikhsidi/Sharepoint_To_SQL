# Importing Modules and Dependencies 
from shareplum import Site
from shareplum import Office365
from sqlalchemy import create_engine, event
import sql_connect as secret
import pyodbc
import pandas as pd
import encrypt as E
from sys import argv
# from deleteSPList import deleteListFunc
from timeit import default_timer as timer
import time

start = timer()

# Decoding the Sharepoint and SQL passwords
pwd = E.decrypt(secret.sharepoint_password.encode("utf-8"))
pwd1 = E.decrypt(secret.password.encode("utf-8"))
# Printing to the console for easy debugging.
print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()

def MetaData():
    '''This function retreive metadat from the admin sharepoint list where there we specify the source SQL table and destination sharepoint list'''
    # Connecting to the Sharepoint site where the Sharepoint_Admin lists lives.
    site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/bidash/', authcookie=authcookie)
    # Reading The Sharepoiny admin input
    # sp_list = site1.List('Adm SqlToSharepoint')
    # Reading the last records in the SQltoshqrepoint list (one sql table at time )
    # data = sp_list.GetListItems('All Items')
    # return data


# Cleaning "None" values from the data Coming from SQL
def dict_clean(d):
    ''' This function clean None values to be empty '''
    for key, value in d.items():
        if value is None:
            value = ''
        d[key] = value
    return d

def SQlTble(data):
    '''This Function retreive data from SQL table and covert it to python object, it takes one object argument containig the source schema, database(db), table '''    
    # Reading All the Values from the sharePoint List.
    _schema = data['SqlSchema']
    _Table = data['SqlTableName']
    db = data['SqlDB']
    
    # connecting to SQL and reading table into a python object
    engine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/{db}?driver=SQL+Server+Native+Client+11.0")
    # df = pd.read_sql_table(_Table, engine, schema=_schema)
    df = pd.read_sql_query(f'select DataSourceSk, code, FRPSCode, Name, FRPSName from {_schema}.{_Table} where DataSourceSk <> 22', engine) #where DataSourceSk <> 22'
    mydata = df.to_dict(orient='records')
    for d in mydata:
        d = dict_clean(d)
        if 'Name' in d.keys():
            d['Name_'] = d.pop('Name')
        if 'Level' in d.keys():
            d['Level_'] = d.pop('Level')
        if 'TypeOfBusinessCode' in d.keys():
            d['TypeOfBusiness'] = d.pop('TypeOfBusinessCode')
    
    return mydata

def pushToSP(data, mydata):
    '''This function push python object to SharePoint lis, its takes tow argumnets both should dictionaries(object) the first one should containe the target site and list, the second one should contain the data to be pushed to the list'''
    _SP = data['SP_Site']
    _SPL = data['SP_List']
    
    print(f'updating Sharepoint List... : {_SPL}....')
    # Connecting to the destination sharepoint list with try
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    except Exception as e : print(e)

    # reading the the desitination Sharepoint list
    mylist1 = site.List(_SPL)
    # # # Adding the new Data to the sharepoint list if data is more than 20000, break it down to batches
    if len(mydata)>5000:
        n=0
        j = 5000
        print("Starting batches ...........")
        while len(mydata)> 0 :
            chunk = mydata[n:j]
            mylist1.UpdateListItems(data=chunk, kind='New')
            print(f"Completed 1st {j} batch-------------")
            n = n + 5000
            j = j + 5000
            time.sleep(60)
    else :
        mylist1.UpdateListItems(data=mydata, kind='New')
    print(f'---------------Done -------------')
    
Total = []


# Fill in the below data object the source Sql table, and the destination list and site.
data = {
    'SqlSchema' : 'frps',
    'SqlTableName' : 'employee_map',
    'SqlDB' : 'frp_edw',
    'SP_Site': 'bidash/direports',
    'SP_List': 'Employee_map'
}

# Pushing to to SP
# for data in MetaData():
# try:
mydata = SQlTble(data)
Total.append(len(mydata))
# deleteListFunc(data)
pushToSP(data, mydata)
# except Exception as e : print(e)
end = timer()
tm = end - start
print (f"Time elapsed to ingest {Total} items into Sharepoint List is : {tm} seconds")
  
