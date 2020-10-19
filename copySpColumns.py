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

start = timer()

# Decoding the Sharepoint and SQL passwords
pwd = E.decrypt(secret.sharepoint_password.encode("utf-8"))
pwd1 = E.decrypt(secret.password.encode("utf-8"))
# Printing to the console for easy debugging.
print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()
# Connecting to the Sharepoint site where the Sharepoint_Admin lists lives.
# li = 'Map_FRPS_Company'
def ListsTobePushed(c, _SPL):
    site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/bidash', authcookie=authcookie)
    # Reading The Sharepoiny admin input
    sp_list = site1.List(_SPL)
    # Reading the last records in the SQltoshqrepoint list (one sql table at time )
    data = sp_list.GetListItems('All Items')
    # retreiving the header of the list
    if type(data) is list:
        return data
    else:      
        header = sp_list.GetListItems('All Items', rowlimit=1)
        # getting the list of fields to be inserted to SQL
        # fields = list(header[0].keys())
        # retreiving the first ID
        i = int(header[0]['ID'])
        n = i + 5000
        m = 5000
        data = []
        # looping over the 5000 chuncks at time because of the limit of SharePoint
        while m == 5000:
            query = {'Where': ['And', ('Geq', 'ID', str(i)), ('Lt', 'ID', str(n))]}
            dt = sp_list.GetListItems(viewname='All Items', query=query) 
            df = pd.DataFrame(dt)
            data_c = df.to_dict('records')
            # data_c = [{k: v for k, v in mydict.items() if k in (c, 'ID')} for mydict in data_c]
            print(data_c[:2])
            data.extend(data_c)
            i = i + 5000
            n = n + 5000
            m = len(dt)
        return data

# Cleaning "None" values from the data Coming from SQL
def dict_clean(d):
    ''' This function clean None values to be empty '''
    for key, value in d.items():
        if value is None:
            value = ''
        d[key] = value
    return d

    # Connecting to the destination sharepoint list with try
def pushToSP(_SP, _SPL, source_c, dest_c): 
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    except Exception as e : print(e)
    mydata = ListsTobePushed(source_c, _SPL)
    for d in mydata:
        # d = dict_clean(d)
        # d.get(dest_c, "empty")
        # d[dest_c] = d.pop(source_c, None)
        d[dest_c] = d.get(source_c, None)
    print(f"After Renaming my column \n {mydata[:10]}")      
    # reading the the desitination Sharepoint list
    mylist1 = site.List(_SPL)
    # Updating the column
    mylist1.UpdateListItems(data=mydata, kind='Update')
    print(f'---------------Done -------------')

# Copying a list items from one site to another
# pushToSP('bidash/direports', 'Map_FRPS_Broker', ListsTobePushed('Map_FRPS_Broker'))

_SP = 'bidash/direports'
_SPL = 'Map_FRPS_Company'
source_c = 'FRPSCode'
dest_c = 'test'

pushToSP(_SP, _SPL, source_c, dest_c)
end = timer()
tm = end - start
print (f"Time elapsed to copy {source_c} into {dest_c} is : {tm} seconds")   