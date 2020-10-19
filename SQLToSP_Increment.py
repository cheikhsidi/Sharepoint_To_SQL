# Importing Modules and Dependencies 
from shareplum import Site
from shareplum import Office365
from sqlalchemy import create_engine, event
import sql_connect as secret
import pyodbc
import pandas as pd
import encrypt as E
from sys import argv
from timeit import default_timer as timer
import time

# Starting the timer 
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
    sp_list = site1.List('Adm SqlToSharepoint')
    # Reading the last records in the SQltoshqrepoint list (one sql table at time )
    data = sp_list.GetListItems('All Items')
    return data

# Cleaning "None" values from the data Coming from SQL
def dict_clean(d):
    ''' This function clean None values to be empty '''
    for key, value in d.items():
        if value is None:
            value = ''
        d[key] = value
    return d


def listObject(mylist1):
    header = mylist1.GetListItems('All Items', rowlimit=1)
    i = int(header[0]['ID'])
    n = i + 5000
    m = 5000
    dt_L = []
    # looping over the 5000 chuncks at time because of the limit of SharePoint
    while m == 5000:
        query = {'Where': ['And', ('Geq', 'ID', str(i)), ('Lt', 'ID', str(n))]}
        dt_ = mylist1.GetListItems(viewname='All Items', query=query)
        dt_L.extend(dt_)
        i = i + 5000
        n = n + 5000
        m = len(dt_)
    return dt_L

def GenID(mylist1, col,  val):
      
    query = {'Where': [('Eq', col, val)]}
    dt = mylist1.GetListItems(viewname='All Items', query=query, rowlimit=1)
    if type(dt) is list:
        print(dt['ID'])
        return dt['ID']
    else :
        dt_L = listObject(mylist1)
        id_val = [d['ID'] for d in dt_L if d[col] == val][0]
        print(id_val)
        return id_val  
     
def SQlTble(data, op):
    '''This Function retreive data from SQL table and covert it to python object, it takes one object argument containig the source schema, database(db), table '''    
    # Reading All the Values from the sharePoint List.
    _schema = data['SqlSchema']
    _Table = data['SqlTableName']
    db = data['SqlDB']  
    
    # connecting to SQL and reading table into a python object
    engine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/{db}?driver=SQL+Server+Native+Client+11.0")
    # df = pd.read_sql_table(_Table, engine, schema=_schema)
    df = pd.read_sql_query(f'select CompanyID, Code, Name from {_schema}.{_Table}_{op}', engine) #where DataSourceSk <> 22'
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

def DatPush(mylist1, mydata, opr):
    if len(mydata)>5000:
        n=0
        j = 5000
        print("Starting batches ...........")
        while len(mydata)> 0 :
            chunk = mydata[n:j]
            mylist1.UpdateListItems(data=chunk, kind=opr)
            print(f"Completed 1st {j} batch-------------")
            n = n + 5000
            j = j + 5000
            time.sleep(60)
    elif len(mydata)>0 :
        mylist1.UpdateListItems(data=mydata, kind=opr)
        print(f" {opr } {len(mydata)} records")
    
    else:
        print(f"No {opr} records")
    
def incrementToSP(data, mydata):
    '''This function push python object to SharePoint lis, its takes tow argumnets both should dictionaries(object) the first one should containe the target site and list, the second one should contain the data to be pushed to the list'''
    _SP = data['SP_Site']
    _SPL = data['SP_List']
    col = data['Identity']
    
    print(f'Inserting new Increments to Sharepoint List... : {_SPL}....')
    # Connecting to the destination sharepoint list with try
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    except Exception as e : print(e)

    # reading the the desitination Sharepoint list
    mylist1 = site.List(_SPL)
    # # # Adding the new Data to the sharepoint list if data is more than 20000, break it down to batches
    new_ids = [d[col] for d in mydata]
    all_Sp_ids = [d[col] for d in listObject(mylist1)]
    # Checking if new increment already being added
    if  set(new_ids).issubset(set(all_Sp_ids)) :
        print("No New increment")
    else :
        DatPush(mylist1, mydata, 'New')
    
def UpdatesToSP(data, mydata):
    '''This function push python object to SharePoint lis, its takes tow argumnets both should dictionaries(object) the first one should containe the target site and list, the second one should contain the data to be pushed to the list'''
    _SP = data['SP_Site']
    _SPL = data['SP_List']
    col = data['Identity']
    
    print(f'Adding updated records to Sharepoint List... : {_SPL}....')
    # Connecting to the destination sharepoint list with try
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    except Exception as e : print(e)

    # reading the the desitination Sharepoint list
    mylist1 = site.List(_SPL)
    # # # Adding the new Data to the sharepoint list if data is more than 20000, break it down to batches
    ks = list(mydata[0].keys())
    for i, d in enumerate(ks):
        if 'Name' in d:
            ks[i] = 'Name_'
    
    all_d = [set({k: d[k] for k in ks}.items()) for d in listObject(mylist1)]
    print(type(all_d))
    print(all_d[0])
    mydata_d = [set(d.items()) for d in mydata]
    print(type(mydata_d))
    print(mydata_d[0])
    if set(mydata_d).issubset(set(all_d)):
        print("No Updates")
    else:
        for d in mydata:
            val = d[col]
            d['ID'] = GenID(mylist1, col,  val)
        DatPush(mylist1, mydata, 'Update')

# Looping through all entries in the admin sharepoint list to update (Add increments and apply updates) sharepoint list from SQL.
for data in MetaData():
    # try:
    st = timer()
    incr = SQlTble(data, 'Increment')
    print(len(incr)) #4
    updt = SQlTble(data, 'Update')
    print(len(updt)) # 60
    incrementToSP(data, incr)
    UpdatesToSP(data, updt)
    ed = timer()
    t = ed - st
    print (f"Time elapsed to incrementing {len(incr)} items and updating {len(updt)} items into Sharepoint List including deleting all records is : {t} seconds")
    # except Exception as e : print(e)
end = timer()
tm = end - start
print (f"Total Time elapsed is : {tm} seconds")
print(f'---------------Done -------------')