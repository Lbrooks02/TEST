from SQLUploader import *

os.system('cls' if os.name == 'nt' else 'clear')
print("------------------------------------------------------------------------------------")

parentFolder = r"C:\Users\young\Desktop\Career\Programming\Projects\SQL-Uploader"

dataTypes = {
    'month_name': 'str',
    'month_nbr': 'int'
    }


SQL_Upload(sourceFile = fr"{parentFolder}\Months.accdb",
           serverName = fr"(LocalDb)\LBrooksServer", 
           databaseName = f"LBrooksDB", 
           tableName = "Months_ACCDB", 
           convertColumnTypes = dataTypes, 
           datasetName="tblMonths"
          )



dataTypes = {
    'MONTH': 'str',
    'NBR': 'int', 
    'PASSED': 'bool'
    }

SQL_Upload(sourceFile = fr"{parentFolder}\Months.xlsx",
           serverName = fr"(LocalDb)\LBrooksServer", 
           databaseName = f"LBrooksDB", 
           tableName = "Months_XLSX", 
           convertColumnTypes = dataTypes
          )