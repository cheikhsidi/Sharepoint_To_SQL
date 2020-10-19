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
from timeit import default_timer as timer

start = timer()

# Decoding the Sharepoint and SQL passwords
pwd = secret.sharepoint_password
pwd1 = secret.password
# Printing to the console for easy debugging.
print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()
# Connecting to the Sharepoint site where the Sharepoint_Admin lists lives.
def deepLokkup (Sp, df1, l2, col, col1, col2):
    site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{Sp}/',  authcookie=authcookie, huge_tree=True)
    sp_list = site1.List(l2)
    data = sp_list.GetListItems('All Items')
    df2 =  pd.DataFrame(data)
    df2 = df2.drop_duplicates(subset=col2)
    df2[col1] =  df2[col1].astype(np.int64).astype(str)
    merge = pd.merge(df1, df2[[col1, col2]], how='left', left_on = col , right_on= col2, validate='m:1', suffixes=('', '_y'))
    merge[col] = np.where(pd.notnull(merge[col1]), merge[col1].astype(str).str.cat(merge[col2],sep=";#"), merge[col])
    merge = merge.replace(np.nan, '', regex=True)
    return merge[list(df1.columns)].to_dict('records')

def lookupFormat(st):
    '''Function to format the lookupfields'''
    if '-' in st:
        d = st.split('-')
        s = f'{d[0]};#{d[1]}'
        return s
    else :
        return ''

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
    print(data[30:34])
    return data
    
    
def ListsTobePushed(Sp, li):
    site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{Sp}/', authcookie=authcookie, huge_tree=True)
    # Reading The Sharepoiny admin input
    sp_list = site1.List(li)
    # Reading the last records in the SQltoshqrepoint list (one sql table at time )
    data = sp_list.GetListItems('All Items')
    if type(data) is list:
        # print(len(data))
        print("this is small list...getting data.....")
        # lookupFields = ['FRPS_DataSourceName']
        # for item in data:
        #     for it in lookupFields:
        #         item[it]= lookupFormat(item[it])
        return data
    else :
        # retreiving the header of the list
        print("large list skipping the try...pushing...")
        return largeList(sp_list)

    # Connecting to the destination sharepoint list with try
def pushToSP(_SP, _SPL, mydata): 
    # try:
    site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    # except Exception as e : print(e)

    # reading the the desitination Sharepoint list
    mylist1 = site.List(_SPL)
    if len(mydata)>20000:
        n=0
        j = 20000
        print("Starting batches ...........")
        while len(mydata)> 0 :
            chunk = mydata[n:j]
            mylist1.UpdateListItems(data=chunk, kind='New')
            print(f"Completed 1st {j} batch-------------")
            n = n + 20000
            j = j + 20000
    else :
        mylist1.UpdateListItems(data=mydata, kind='New')
    print(f'---------------Done -------------')

# Copying a list items from one site to another
# lists = ['Map_FRPS_AgencyBranchProfitCenter_']
# for l in lists:
pushToSP('Technology-Data', 'DataSource', ListsTobePushed('bidash', 'FRPS DataSource'))

# When there is field need to be converted to a lookup field.
# Sp = 'bidash'
# df1 = pd.DataFrame(ListsTobePushed('bidash', 'Map_FRPS_AgencyBranchProfitCenter'))
# pushToSP('bidash', 'Map_FRPS_AgencyBranchProfitCenter_',deepLokkup(Sp, df1, 'FRPS DataSource', 'FRPS_DataSourceName', 'SK', 'Name_'))
# Sp = 'bidash'

# final = deepLokkup(Sp, df1, 'FRPS DataSource', 'FRPS_DataSourceName', 'SK', 'Name_')
# print(l)
end = timer()
tm = end - start
print (f"Time elapsed to copy items is : {tm/60} minutes")   
