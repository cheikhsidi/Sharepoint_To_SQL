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

# Decoding the Sharepoint and SQL passwords
pwd = secret.sharepoint_password
pwd1 = secret.password
# Printing to the console for easy debugging.
print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()

def SPList(_SP, _SPL):
    try:
        site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
        mylist1 = site.List(_SPL)
        data = mylist1.GetListItems('All Items')
        return data
    except Exception as e : print(e)

df1 = pd.read_excel('../../../Downloads/User Identity Final Version (1).xlsx').replace(np.nan, '', regex=True) 

def deepLokkup (Sp, df1, l2, col, col1, col2):
    site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{Sp}/', authcookie=authcookie, huge_tree=True)
    sp_list = site1.List(l2)
    data = sp_list.GetListItems('All Items')
    df2 =  pd.DataFrame(data)
    df2 = df2.drop_duplicates(subset=col2)
    df2[col1] =  df2[col1].astype(np.int64).astype(str)
    merge = pd.merge(df1, df2[[col1, col2]], how='left', left_on = col , right_on= col2, validate='m:1', suffixes=('', '_y'))
    merge[col] = np.where(pd.notnull(merge[col1]), merge[col1].astype(str).str.cat(merge[col2],sep=";#"), merge[col])
    merge = merge.replace(np.nan, '', regex=True)
    # print(merge[list(df1.columns)].to_dict('records')[0])
    return merge[list(df1.columns)]
# def lookupFormat(st):
#     '''Function to format the lookupfields'''
#     if '-' in st:
#         d = st.split('-')
#         s = f'{d[0]};#{d[1]}'
#         return s
#     else :
#         return ''
col1 = 'Source System'
col2 = 'Business Unit'
col3 = 'Deal Name'


list1 = list(df1.columns)
cols = [i for i in list1 if i not in [col2, col3]]

col1_list = []
df2 = deepLokkup ('bidash', df1, 'FRPS DataSource', 'Source System', 'SK', 'Name_')[cols]
df3 = deepLokkup ('bidash', df1, 'FRPS Business Unit', 'Business Unit', 'ID', 'Name_')[[col2]]
df4 = deepLokkup ('bidash', df1, 'FRPS Deal', 'Deal Name', 'ID', 'DealName')[[col3]]


# col1 = 'Source System'
# col2 = 'Business Unit'
# col3 = 'Deal Name'
list1 = list(df2.columns)
cols = [i for i in list1 if i not in [col1, col2, col3]]

# df3 = df3.drop_duplicates(subset=[col1, col2, col3])
# df4 = df4.drop_duplicates(subset=[col1, col2, col3])
# final = df2.join(df3, how='left')\
#     .join(df4, how = 'left')

final = pd.concat([df2, df3, df4], axis=1, sort=False)
# final.drop(['Source System_x', 'Business Unit', 'Deal Name', 'Deal Name_x', 'Source System_y', 'Business Unit_y'], axis=1, inplace=True)
# final.rename(columns={'Business Unit_x':'Business Unit', 'Deal Name_y':'Deal Name'}, inplace=True)
print(final.to_dict('records')[0])
mydata = final.replace(np.nan, '', regex=True).to_dict('records')

    
    
    
    
_SP = 'bidash'
_SPL = 'FRPS User Identity'

print(f'updating Sharepoint List... : {_SPL}....')
# Connecting to the destination sharepoint list with try 
print(mydata[0])
# try:
site = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/{_SP}/', authcookie=authcookie)
# except Exception as e : print(e)
mylist1 = site.List(_SPL)
data1 = mylist1.GetListItems('All Items')

# lists = [item['Title'] for item in site.GetListCollection()]

# Adding the new Data to the sharepoint list 
mylist1.UpdateListItems(data=mydata, kind='New')
print(f'---------------Done -------------')






