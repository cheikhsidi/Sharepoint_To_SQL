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

def Lists(Sp):
    site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{Sp}/', authcookie=authcookie, huge_tree=True)
    # mylist = site.List(li)
    li = site.GetListCollection()
    # df = pd.DataFrame(li)
    # df.to_excel('bidash_lists.xlsx', index=False)
    mylists = [l['Title'] for l in li]
    secur = {l['Title']:l['InheritedSecurity'] for l in li}
    readSecurity = {l['Title']:l['ReadSecurity'] for l in li}
    Allowance = {l['Title']:l['AllowAnonymousAccess'] for l in li} 
    
    depends = {}
    
    cols = ['Business', 'Operating', 'Deal', 'DataSource']
    
    # Get all list that have dependencies
    for lis in mylists :
        temp = []
        try:
            mylist = site.List(lis)
            data = mylist.GetListItems('All Items', rowlimit=1)[0]
            for col in cols :
                if any(col in s for s in data.keys()):
                    temp.append(col)
            if len(temp) != 0 :
                depends[lis] = temp
        except Exception as e:
            print(e)
    print(depends)  
    # print(mylists)
    # print(secur)
    # print(readSecurity)
    # print(Allowance)
    # print(f'\n{li}')
    
Lists('bidash')