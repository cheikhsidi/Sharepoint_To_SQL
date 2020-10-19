# Importing Modules and Dependencies 
from shareplum import Site
from shareplum import Office365
from sqlalchemy import create_engine, event
import sql_connect as secret
import pyodbc
import pandas as pd
import encrypt as E
import numpy as np
from sys import argv
from timeit import default_timer as timer
import time
from datetime import datetime

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
    global sp_list
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

def sp_conn(data):
    _SP = data['SP_Site']
    _SPL = data['SP_List']
    
    # Connecting to the destination sharepoint list with try
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    except Exception as e : print(e)

    # reading the the desitination Sharepoint list
    mylist1 = site.List(_SPL)
    return mylist1
def listObject(data):
    '''This function takes a list object as an argument and rturn a dataframe of the sharpoint list to be updated'''
    mylist1 = sp_conn(data)
    d = mylist1.GetListItems('All Items')
    if type(d) is list:
        return pd.DataFrame(d)
    else:
        header = mylist1.GetListItems('All Items', rowlimit=1)
        i = int(header[0]['ID'])
        n = i + 5000
        m = 5000
        dt_L = []
        # looping over the 5000 chuncks at time because of the limit of SharePoint
        while m == 5000:
            try: 
                query = {'Where': ['And', ('Geq', 'ID', str(i)), ('Lt', 'ID', str(n))]}
                dt_ = mylist1.GetListItems(viewname='All Items', query=query)
                dt_L.extend(dt_)
                i = i + 5000
                n = n + 5000
                m = len(dt_)
            except Exception as e : print(e)
        list_df = pd.DataFrame(dt_L)
        return list_df
     
def SQlTble(data, op, lastRun, list_col):
    '''This Function retreive data from SQL table and covert it to python object, it takes three object arguments containig the source schema, database(db), table and the operation type (increment or updates) and the last Run datatime'''    
    # Reading All the Values from the sharePoint List.
    _schema = data['SqlSchema']
    _Table = data['SqlTableName']
    db = data['SqlDB']  
    
    # connecting to SQL and reading table into a python object
    engine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/{db}?driver=SQL+Server+Native+Client+11.0")
    lastConst = f"convert(datetime,convert(nvarchar,'{str(lastRun)}',1))"
    t = f"{_schema}.{_Table}"
    Increment = f"CONVERT(datetime, SWITCHOFFSET(BatchInsertDate, DATEPART(TZOFFSET, BatchInsertDate AT TIME ZONE 'Eastern Standard Time')))"
    Updates = f"CONVERT(datetime, SWITCHOFFSET(BatchUpdateDate, DATEPART(TZOFFSET, BatchUpdateDate AT TIME ZONE 'Eastern Standard Time')))"
    queryIncrement = f"select max({Increment}) as BID from {t}"
    queryUpdates = f"select max({Updates}) as BUD from {t}"
        
    # Retreiving Last Insert date
    BID = pd.read_sql_query(queryIncrement, engine)
    SQLBatchInsertDate = BID.to_dict(orient='records')[0]
    # Retreiving Last Update date
    BUD = pd.read_sql_query(queryUpdates, engine)
    SQLBatchUpdateDate = BUD.to_dict(orient='records')[0]
    # Exposing Insert and Upadate dates
    global SQL_BID, SQL_BUD
    SQL_BID = SQLBatchInsertDate['BID']
    SQL_BUD = SQLBatchUpdateDate['BUD']
    
    # Reading the data from the SQL
    if op.lower() == 'increment':
        df = pd.read_sql_query(f"select * from {t} where {Increment} > {lastConst}", engine)
    else :
        df = pd.read_sql_query(f"select * from {t} where {Updates} > {lastConst}", engine)
    df = df[np.intersect1d(df.columns,  list_col)]    
    mydata = df.to_dict(orient='records')
    for d in mydata:
        d = dict_clean(d)
        if 'Name' in d.keys():
            d['Name_'] = d.pop('Name')
        if 'Level' in d.keys():
            d['Level_'] = d.pop('Level')
        if 'TypeOfBusinessCode' in d.keys():
            d['TypeOfBusiness'] = d.pop('TypeOfBusinessCode')
        if 'Modified By' in d.keys():
            del d['Modified By']
    return mydata

def DatPush(mylist1, mydata, opr):
    '''This function is the Meat function to actually performs the push to the sharepoint it takes three arguments the target liast object, the data aboject to be ingested, and the type of operation increment or updates'''
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
            print(f" Large list, {len(mydata)}  { opr } records")
    elif len(mydata)>0 :
        mylist1.UpdateListItems(data=mydata, kind=opr)
        print(f" {opr } {len(mydata)} records")
    
    else:
        print(f"No {opr} records")

def UpdateLastRunDate(sp_list, d):
    '''This function updates the last run datetime field in the admin list to keep track of the increments and updates'''
    sp_list.UpdateListItems(data=d, kind='Update')
   
def incrementToSP(data, mydata):
    '''This function push python object to SharePoint list, it takes tow argumnets both should dictionaries(object) 
    the first one should containe the target site and list, the second one should contain the data to be pushed to the list'''
    # reading the the desitination Sharepoint list
    mylist1 = sp_conn(data)
    # # # Adding the new Data to the sharepoint list if data is more than 20000, break it down to batches
    DatPush(mylist1, mydata, 'New')
    
def UpdatesToSP(data, mydata):
    '''This function update SharePoint list from the Sql object, it takes tow argumnets both should dictionaries(objects) the first one should containe the target site, the list and the Identity column, 
    the second one should contain the sql data object to be pushed to the list'''
    col = data['Identity']
    
    print(f'updating {len(mydata)} records to Sharepoint List....')
    mylist1 = sp_conn(data)
     # Retreiving the ids for teh updated records
    #convert mydata object to a datframe
    dtf = pd.DataFrame(mydata)
    #convert the id column to a list
    id_col = dtf[col].tolist()
    #get the the dataframe of the entire list
    df = listObject(data)
    #filter the dataframe down to the list of new change
    df1 = df[df[col].isin(id_col)]
    #pull the ids of the records
    ids = df1['ID'].tolist()
    #adding the ids as a the ID column to the new change to be ingested
    dtf['ID'] = ids
    print(dtf.head())
    # Covert it back to a python pbject
    mydata = dtf.to_dict(orient='records')
    #update the list
    DatPush(mylist1, mydata, 'Update')

# Looping through all entries in the admin sharepoint list to update (Add increments and apply updates) sharepoint list from SQL.
for data in MetaData():
    # reating meta data from the admin list, and retreiving data from sql tables 
    lastRun = data['RunDateTime']
    st = timer()
    list_col = listObject(data).columns.tolist()
    print(list_col)
    
    incr = SQlTble(data, 'Increment', lastRun, list_col)
    updt = SQlTble(data, 'Update', lastRun, list_col)
    print(incr)
    # checking if there are increments and updates to be applyed teh increments and the updates 
    # try:
    print (f"sql_date insert: {SQL_BID}, updates : {SQL_BUD}, and the last run :{lastRun}")
    if SQL_BID > lastRun:
        incrementToSP(data, incr)
    else :
        print("No Increment")
        
    if SQL_BUD > lastRun:
        UpdatesToSP(data, updt)
        
    else :
        print("No Updates")
    ed = timer()
    t = ed - st
    print (f"Time elapsed to incrementing {len(incr)} items and updating {len(updt)} items into Sharepoint List including deleting all records is : {t} seconds")
    # except Exception as e : print(e)
    
now = datetime.now()
dt = [{'ID':data['ID'], 'RunDateTime': now}]
UpdateLastRunDate(sp_list, dt)
end = timer()
tm = end - start
print (f"Total Time elapsed is : {tm} seconds")
print(f'---------------Done -------------')