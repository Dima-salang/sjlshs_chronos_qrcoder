import pandas as pd
import sqlite3
import os
from datetime import datetime
from google.cloud import firestore
from firebase_client import db

class ImporterBuilder:
    """ Importer builder class """

    def __init__(self, file):
        self.file = file

    def build(self):
        if not os.path.exists(self.file):
            raise ValueError("File does not exist.")
        # return the appropriate importer
        if self.file.endswith('.xlsx') or self.file.endswith('.xls'):
            return ExcelDataImporter(self.file)
        else:
            raise ValueError("Unsupported file format.")

class DataImporter:
    """ Abstract importer class for importing data from various repositories such as Firestore """
    def __init__(self):
        pass
    
    def import_data(self, start_date: datetime, end_date: datetime, section: str | None = None):
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
    
    def import_data(self, start_date: datetime, end_date: datetime, section: str | None = None):
        """Gets all student attendance records from Firestore for the given date range."""
        try:
            query = self.db.collection('attendance').where('timestamp', '>=', start_date).where('timestamp', '<=', end_date)

            if section:
                query = query.where('studentSection', '==', section)
            
            return query.get()
        except Exception as e:
            print(f"Error getting student records: {e}")
            raise

    


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
    
    def import_data(self, start_date: datetime, end_date: datetime, section: str | None = None):
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

    def upload_master_list_to_firestore(self):
        """
        Uploads the master list from the Excel file to Firebase Firestore.
        """
        if db is None:
            print("Error: Firestore client is not initialized.")
            return

        df = self.parse_excel_file()
        batch = db.batch()
        collection_ref = db.collection('master_list')
        
        count = 0
        BATCH_SIZE = 400 # Firestore batch limit is 500

        print(f"Uploading {len(df)} records to Firestore...")

        for index, row in df.iterrows():
            # Create a document with LRN as ID
            doc_ref = collection_ref.document(str(row['LRN']))
            
            student_data = {
                'lrn': str(row['LRN']),
                'last_name': row['LAST_NAME'],
                'first_name': row['FIRST_NAME'],
                'student_year': int(row['STUDENT_YEAR']) if str(row['STUDENT_YEAR']).isdigit() else row['STUDENT_YEAR'],
                'section': row['SECTION'],
                'adviser': row['ADVISER'],
                'gender': row['GENDER']
            }
            
            batch.set(doc_ref, student_data)
            count += 1

            if count >= BATCH_SIZE:
                batch.commit()
                batch = db.batch()
                count = 0
                print("Committed batch.")

        if count > 0:
            batch.commit()
            print("Committed final batch.")
        
        print("Master list upload complete.")




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

class MasterListManager:
    """
    Manages the master list data in the local SQLite database and Firestore.
    """
    def __init__(self, db_path='master_list.db'):
        self.db_path = db_path

    def get_local_records(self):
        """Retrieves all records from the local SQLite database."""
        if not os.path.exists(self.db_path):
            return []
            
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT * FROM master_list")
            columns = [description[0] for description in cursor.description]
            rows = cursor.fetchall()
            
            records = []
            for row in rows:
                records.append(dict(zip(columns, row)))
            return records
        except sqlite3.OperationalError:
            # Table might not exist
            return []
        finally:
            conn.close()

    def upload_local_to_firestore(self, progress_callback=None):
        """Uploads all local records to Firestore."""
        if db is None:
            raise ConnectionError("Firestore client is not initialized.")

        records = self.get_local_records()
        if not records:
            print("No local records to upload.")
            return 0

        batch = db.batch()
        collection_ref = db.collection('master_list')
        count = 0
        BATCH_SIZE = 400

        print(f"Uploading {len(records)} records from local DB to Firestore...")

        for record in records:
            # Ensure data types match what Firestore expects
            doc_ref = collection_ref.document(str(record['lrn']))
            
            # Clean up data if needed (e.g. ensure student_year is int)
            try:
                student_year = int(record['student_year'])
            except (ValueError, TypeError):
                student_year = record['student_year']

            student_data = {
                'lrn': str(record['lrn']),
                'last_name': record['last_name'],
                'first_name': record['first_name'],
                'student_year': student_year,
                'section': record['section'],
                'adviser': record['adviser'],
                'gender': record['gender']
            }
            
            batch.set(doc_ref, student_data)
            count += 1

            if count % BATCH_SIZE == 0:
                batch.commit()
                batch = db.batch()
                if progress_callback:
                    progress_callback(count)
                print(f"Committed batch of {BATCH_SIZE} records.")

        if count % BATCH_SIZE != 0:
            batch.commit()
            if progress_callback:
                progress_callback(count)
            print("Committed final batch.")
            
        return count

    def delete_local_records(self):
        """Deletes all records from the local SQLite database."""
        if not os.path.exists(self.db_path):
            return 0
            
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM master_list")
            deleted = cursor.rowcount
            conn.commit()
            print(f"Deleted {deleted} local records.")
            return deleted
        except sqlite3.OperationalError:
            print("Table master_list does not exist.")
            return 0
        finally:
            conn.close()

    def delete_firestore_records(self, progress_callback=None):
        """Deletes all records from the Firestore master_list collection."""
        if db is None:
            raise ConnectionError("Firestore client is not initialized.")

        collection_ref = db.collection('master_list')
        batch_size = 400
        deleted = 0

        def delete_batch(docs):
            batch = db.batch()
            for doc in docs:
                batch.delete(doc.reference)
            batch.commit()

        while True:
            docs = list(collection_ref.limit(batch_size).stream())
            if not docs:
                break
            
            delete_batch(docs)
            deleted += len(docs)
            if progress_callback:
                progress_callback(deleted)
            print(f"Deleted {len(docs)} records from Firestore.")
            
        return deleted