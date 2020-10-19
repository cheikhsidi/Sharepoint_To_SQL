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
            data = df.to_dict('records')
        else :
            print(f"{count+1} batch...{len(df)}")
            data.extend(df.to_dict('records'))

        i = i + 5000
        n = n + 5000
        m = len(dt)
        count += 1
    print(f"largeList count : {len(data)}")
    return data

def conn(SP, li):
    site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{Sp}/', authcookie=authcookie, huge_tree=True)
    # Reading The Sharepoiny admin input
    sp_list = site1.List(li)
    # Reading the last records in the SQltoshqrepoint list (one sql table at time )
    data = sp_list.GetListItems('All Items')
    return data
  
def customUpdate(obj1, mydata):
    
    
    
    
def ListsTobePushed(Sp, li):
    site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{Sp}/', authcookie=authcookie, huge_tree=True)
    # Reading The Sharepoiny admin input
    sp_list = site1.List(li)
    # Reading the last records in the SQltoshqrepoint list (one sql table at time )
    data = sp_list.GetListItems('All Items')
    if type(data) is list:
        # print(len(data))
        print("this is small list...getting data.....")
        return data
    else :
        # retreiving the header of the list
        print("large list skipping the try...pushing...")
        return largeList(sp_list)
               
# Cleaning "None" values from the data before pushing it
def dict_clean(d):
    ''' This function clean None values to be empty '''
    for key, value in d.items():
        if value is None:
            value = ''
        d[key] = value
    return d

    # Connecting to the destination sharepoint list with try
def pushToSP(_SP, _SPL, mydata): 
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    except Exception as e : print(e)

    # reading the the desitination Sharepoint list
    mylist1 = site.List(_SPL)
    # data1 = mylist1.GetListItems('All Items')
    
    # Retreiving all IDs of the list if it not empty
    # ids = [item['ID'] for item in data1]
    # Delete all items from the sharepoint list by IDs
    # mylist1.UpdateListItems(ids, kind='Delete')
    # # # Adding the new Data to the sharepoint list if data is more than 20000, break it down to batches
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
pushToSP('bidash', 'Key Roles', ListsTobePushed('Technology-Data', 'Key Roles'))
end = timer()
tm = end - start
print (f"Time elapsed to copy items is : {tm} seconds")   
