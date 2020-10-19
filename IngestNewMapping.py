from sqlalchemy import Table, Column, Integer, String, MetaData, ForeignKey
from sqlalchemy.sql.expression import Executable, ClauseElement
from sqlalchemy import create_engine, inspect, event
import sql_connect as secret
from shareplum import Site
# from shareplum.Site import Version
from shareplum import Office365
import pyodbc
import pandas as pd
import os
import fnmatch
from datetime import date
import time
# import encrypt as E

today = date.today()
# YY_mm_dd
d = today.strftime("%Y_%m_%d")
# print("d =", d)
# Decoding the Sharepoint and SQL passwords
pwd = secret.sharepoint_password
pwd1 = secret.password
# Printing to the console for easy debugging.
basepath = '../Integration/Micheletti/'




print ("Starting......................\n")
# setting up the connection and authentication to sharepint site
authcookie = Office365('https://foundationriskpartners.sharepoint.com', username=secret.sharepoint_username, password=pwd).GetCookies()
site1 = Site(f'https://foundationriskpartners.sharepoint.com.us3.cas.ms/sites/bidash/biteam/', authcookie=authcookie, huge_tree=True)


engine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/frp_edw?driver=SQL+Server+Native+Client+11.0")
@event.listens_for(engine, "before_cursor_execute")
def receive_before_cursor_execute(
    conn, cursor, statement, params, context, executemany
        ):
        if executemany:
            cursor.fast_executemany = True             
                    
sheets = ['Companies', 'Brokers', 'Vendors', 'PolicyCodes', 'ActivityCodes', 'LineStatusCodes', 'employee']

def OldNew(basepath, filename):
    # basepath = 'my_directory/'
    for sh in sheets :
        df = pd.read_excel(f'{basepath}{filename}.xlsx', sheet_name=sh)
        df.to_sql(f'OldNew_Samples', engine, index=False, schema="etl", if_exists = 'replace')
        
        sql = 'exec OldNew_SP (?)'
        values = (sh)
        with engine.begin() as conn:
            conn.execute(sql, (values))
        time.sleep(60)
  
                               
def ingest_sample0(basepath):
    # basepath = 'my_directory/'
    for file in os.listdir(basepath) :
        filename, file_extension = os.path.splitext(file)
        print (filename)
        print (file_extension)
        if file_extension == '.csv':
            df = pd.read_csv(f'{basepath}{filename}.csv')
        elif file_extension == '.xlsx':
            df = pd.read_excel(f'{basepath}{filename}.xlsx')  
            
        df.to_sql(f'{filename}_Sample0', engine, index=False, schema="etl", if_exists = 'replace')


