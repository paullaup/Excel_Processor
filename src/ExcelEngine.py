import warnings
warnings.filterwarnings("ignore")
import adodbapi
import pandas

class ExcelEngine:

    #adodbapi.apibase.OperationalError connection error
    def __init__(self, excelPath: str):
        connectionString = rf'Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelPath};Extended Properties="Excel 12.0 Xml;HDR=YES;"'
        self.connection = adodbapi.connect(connectionString)

    def close(self):
        self.connection.close()

    #pandas.errors.DatabaseError
    def readAll(self, sheetName: str) -> pandas.DataFrame:
        sqlQuery = f'''
                    SELECT
                        *
                    FROM
                        [{sheetName}$]
                '''
        return pandas.read_sql(sqlQuery, self.connection)
    
    def groupSum(self, sheetName: str, columnsGroup: list, columnsSum: list) -> pandas.DataFrame:
        columnsSumFunction = [f"SUM({column})" for column in columnsSum]
        sqlquery = f'''
                    SELECT
                        {",".join(columnsGroup)},
                        {",".join(columnsSumFunction)}
                    FROM
                        [{sheetName}$]       
                    '''
        return pandas.read_sql()
    
    
        