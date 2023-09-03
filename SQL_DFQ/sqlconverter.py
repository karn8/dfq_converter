import pyodbc
import pandas as pd
import warnings

# server = 'NT Service\\MSSQLSERVER'  
server = 'ASUS-VIVOBOOK'
database = 'test_db'        
username = 'dbo'           
password = 'database'            

connection_string = rf'DRIVER={{SQL Server}};SERVER={server};DATABASE={database};Trusted_Connection=yes;'

connection = pyodbc.connect(connection_string)

warnings.filterwarnings("ignore", category=UserWarning, message="pandas only supports SQLAlchemy connectable")

query = 'SELECT * FROM Test_db_1'
df = pd.read_sql_query(query, connection)

excel_file_path = 'output.xlsx'
df.to_excel(excel_file_path, index=False)

connection.close()