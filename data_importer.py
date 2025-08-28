import pandas as pd
import sqlite3
from models import StudentList

class DataImporter:
    """ Abstract importer class for importing data from various repositories such as Firestore """
    def __init__(self):
        pass
    
    def import_data(self):
        raise NotImplementedError





class FirestoreDataImporter(DataImporter):
    """ Firestore data importer class 
    
    Args:
        db_client: A valid Firestore database client
    
    """

    def __init__(self, db_client):
        super().__init__()
        if not db_client:
            raise ValueError("A valid Firestore database client is required.")
        self.db = db_client
    
    def import_data(self):
        pass

    


class ExcelDataImporter(DataImporter):
    
    """ Excel data importer class 
    
    Args:
        excel_path: Path to the Excel file
    """







    def __init__(self, excel_path):
        super().__init__()
        if not excel_path:
            raise ValueError("An Excel file path is required.")
        self.excel_path = excel_path
    
    def import_data(self):
        pass


    # import master list of students
    def import_master_list(self):
        """
        Import the master list of students from the Excel file.

        The master list of students is stored in a SQLite database.

        The format of the excel should be as follows:

        LRN, LAST_NAME, FIRST_NAME, STUDENT_YEAR, SECTION, ADVISER, GENDER
        
        """
        df = self.parse_excel_file()
        self.store_master_list(df)

    

    
    # parse excel file
    def parse_excel_file(self):
        df = pd.read_excel(self.excel_path)
        df = df.fillna('')

        valid_columns = ['LRN', 'LAST_NAME', 'FIRST_NAME', 'STUDENT_YEAR', 'SECTION', 'ADVISER', 'GENDER']

        # validate excel file
        if df is None:
            raise ValueError("The Excel file is empty.")
        # validate columns
        if not all(col in df.columns for col in valid_columns):
            raise ValueError("The Excel file is missing one or more required columns.")

        
        return df



    # store master list of students in SQLite database
    def store_master_list(self, df):
        conn = sqlite3.connect('master_list.db')
        self.create_table(conn)

        for index, row in df.iterrows():
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO master_list (lrn, last_name, first_name, student_year, section, adviser, gender)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (row['LRN'], row['LAST_NAME'], row['FIRST_NAME'], row['STUDENT_YEAR'], row['SECTION'], row['ADVISER'], row['GENDER']))
            conn.commit()
        
        conn.close()




    def create_table(self, conn):
        cursor = conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS master_list (
                lrn TEXT PRIMARY KEY,
                last_name TEXT,
                first_name TEXT,
                student_year INT,
                section TEXT,
                adviser TEXT,
                gender TEXT
            )
        ''')
        conn.commit()


        