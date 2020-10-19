from sqlalchemy import create_engine, Table, Column, Integer, Unicode, MetaData, String, Text, update, and_, select, func, types
from sqlalchemy.ext.automap import automap_base
import sql_connect as secret

pwd1 = secret.password
# create engine, reflect existing columns, and create table object for oldTable
#engine_prod = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/{secret.db}?driver=SQL+Server+Native+Client+11.0")
#engine_dev = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/{secret.db}?driver=SQL+Server+Native+Client+11.0")
# tables = ['frps.Deal', 'frps.DataSource', 'frps.Chart_Agnecy', 'frps.BusinessUnit', 'frps.OperatingUnit', 'frps.ReportableUnit']


# create engine, reflect existing columns, and create table object for oldTable
srcEngine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/frpbi?driver=SQL+Server+Native+Client+11.0")
srcEngine_metadata = MetaData(bind=srcEngine)
# srcEngine_metadata.reflect(srcEngine) # get columns from existing table
srcTable = Table('ExternalUser_Profile', srcEngine_metadata, autoload=True, schema="etl")

# # create engine and table object for newTable
destEngine = create_engine(f"mssql+pyodbc://{secret.user}:{pwd1}@{secret.server}:1433/frp_edw?driver=SQL+Server+Native+Client+11.0")
destEngine_metadata = MetaData(bind=destEngine)
destTable = Table('ExternalUser_Profile', destEngine_metadata, schema="etl")


# copy schema and create newTable from oldTable
for column in srcTable.columns:
    destTable.append_column(column.copy())
destTable.create()
    
# Copy Data from the source table to the new table
insert = destTable.insert()
for row in srcTable.select().execute():
    insert.execute(row)