from sqlalchemy import Table, Column, Integer, String, MetaData, ForeignKey
from sqlalchemy.sql.expression import Executable, ClauseElement
from sqlalchemy import create_engine, inspect, event
import sql_connect as secret
import pyodbc
import pandas as pd
# import encrypt as E

pwd1 = secret.password
# tables = ['OperatingUnit']
df = pd.read_excel('../Integration/Micheletti/Micheletti_Sample0.xlsx', skiprows=1, sheet_name='PolicyLineType')

engine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/frp_edw?driver=SQL+Server+Native+Client+11.0")

@event.listens_for(engine, "before_cursor_execute")
def receive_before_cursor_execute(
       conn, cursor, statement, params, context, executemany
        ):
            if executemany:
                cursor.fast_executemany = True

df.to_sql('Micheletti_S0', engine, index=False, schema="etl", if_exists = 'replace')

# df.to_excel("HIPI_CustomReport.xlsx", index=False)
# table = Table('frps.{}', metadata, autoload=True, autoload_with=db1)



# table.create(engine=db2)

