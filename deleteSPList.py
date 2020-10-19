# Importing Modules and Dependencies 
from shareplum import Site
from shareplum import Office365
from sqlalchemy import create_engine
import sql_connect as secret
import pyodbc
import pandas as pd
# import encrypt as E
import time
from timeit import default_timer as timer

start = timer()
# Decoding the Sharepoint and SQL passwords
pwd = secret.sharepoint_password
# Printing to the console for easy debugging.
print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()

def deleteListFunc(_SP, _SPL):
    '''This function truncate Sharepoint list, it takes two argumentsthe target site and list to be deleted'''
    print(f'deleting items in the Sharepoint List : {_SPL}....')
    # Connecting to the destination sharepoint list with try 
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
    except Exception as e : print(e)

    # reading the the desitination Sharepoint list items
    mylist1 = site.List(_SPL)
    i = 1
    count=0
    while i > 0:
        # lopping while there are items
        data1 = mylist1.GetListItems('All Items', rowlimit=2000)
        ids = [item['ID'] for item in data1]
        # Delete all selected items from the sharepoint list by IDs
        mylist1.UpdateListItems(ids, kind='Delete')
        if len(data1) == 2000 :
            print(f"deleted {count + 2000} chunck...")
            time.sleep(300) 
        i = len(data1)
        count = count + 2000
        
_SP = 'bidash' # the Site where the target list lives
_SPL = 'zzzGlossary Sharepoint Lists' # the targer list name to be deleted
# li = ['zzzChartAgency', 'zzzBackup_reportable', 'zzz_FRPS Chart Agency', 'zzFRPS Reportable_Units']
deleteListFunc(_SP, _SPL)
# for l in li:
#     try:
#         deleteListFunc(_SP, l)
#     except Exception as e : print(e)

end = timer()
tm = end - start
print (f"Time elapsed to delete items is : {tm/60} minutes")     
print("-----------Done-------------")


