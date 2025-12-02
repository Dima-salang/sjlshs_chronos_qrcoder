import unittest
from unittest.mock import MagicMock, patch
import os
import pandas as pd
import shutil
from data_importer import ExcelDataImporter
from image_manager import DriveImageManager as ImageManager
from key_manager import KeyManager
from qr_generator import QRCodeGenerator

class TestAutomation(unittest.TestCase):

    def setUp(self):
        # Create dummy directories and files
        os.makedirs("test_excel", exist_ok=True)
        os.makedirs("test_images/Section A", exist_ok=True)
        os.makedirs("test_qr", exist_ok=True)
        
        # Create dummy Excel
        self.excel_path = "test_excel/masterlist.xlsx"
        df = pd.DataFrame({
            'LRN': ['123456789012'],
            'LAST_NAME': ['Doe'],
            'FIRST_NAME': ['John'],
            'STUDENT_YEAR': [12],
            'SECTION': ['Section A'],
            'ADVISER': ['Mr. Smith'],
            'GENDER': ['M'],
            'Student ID': ['123456789012'], # For QR generator
            'Section': ['Section A']      # For QR generator
        })
        df.to_excel(self.excel_path, index=False)

        # Create dummy image
        with open("test_images/Section A/student.png", "w") as f:
            f.write("dummy image content")

        # Create dummy key
        with open("test_key.key", "w") as f:
            f.write("dummy key content")

    def tearDown(self):
        # Cleanup
        shutil.rmtree("test_excel")
        shutil.rmtree("test_images")
        shutil.rmtree("test_qr")
        if os.path.exists("test_key.key"):
            os.remove("test_key.key")

    @patch('data_importer.db')
    def test_masterlist_upload(self, mock_db):
        print("\nTesting Masterlist Upload...")
        mock_batch = MagicMock()
        mock_db.batch.return_value = mock_batch
        mock_collection = MagicMock()
        mock_db.collection.return_value = mock_collection

        importer = ExcelDataImporter(self.excel_path)
        importer.upload_master_list_to_firestore()

        mock_db.collection.assert_called_with('master_list')
        mock_batch.set.assert_called()
        mock_batch.commit.assert_called()
        print("Masterlist upload test passed.")

    @patch('image_manager.build')
    @patch('image_manager.service_account.Credentials')
    def test_image_upload(self, mock_creds, mock_build):
        print("\nTesting Image Upload (Google Drive)...")
        
        # Mock Drive Service
        mock_service = MagicMock()
        mock_build.return_value = mock_service
        
        # Mock Files Resource
        mock_files = MagicMock()
        mock_service.files.return_value = mock_files
        
        # Mock List (Search) - Return empty list first (folder doesn't exist) then folder id
        # We need to handle multiple calls to list()
        # 1. Check root folder -> returns [] (so it creates it)
        # 2. Check subfolder -> returns [] (so it creates it)
        # 3. Check file -> returns [] (so it uploads it)
        
        # Simulating responses for list()
        # We can just return empty lists to force creation/upload for all
        mock_list = MagicMock()
        mock_list.execute.return_value = {'files': []}
        mock_files.list.return_value = mock_list
        
        # Mock Create
        mock_create = MagicMock()
        mock_create.execute.return_value = {'id': 'dummy_id'}
        mock_files.create.return_value = mock_create

        # Create dummy credentials file
        with open("credentials.json", "w") as f:
            f.write("{}")

        try:
            manager = ImageManager("test_images")
            manager.upload_images()

            # Verify build was called (auth)
            mock_build.assert_called()
            
            # Verify create was called (folders + file)
            # We expect at least one file upload
            self.assertTrue(mock_files.create.called)
            print("Image upload test passed.")
        finally:
            if os.path.exists("credentials.json"):
                os.remove("credentials.json")

    @patch('key_manager.db')
    def test_key_upload(self, mock_db):
        print("\nTesting Key Upload (Firestore)...")
        mock_doc_ref = MagicMock()
        mock_db.collection.return_value.document.return_value = mock_doc_ref

        manager = KeyManager("test_key.key")
        manager.upload_key()

        mock_db.collection.assert_called_with('config')
        mock_doc_ref.set.assert_called()
        # Check if set was called with a dict containing 'encryption_key'
        args, kwargs = mock_doc_ref.set.call_args
        self.assertIn('encryption_key', args[0])
        print("Key upload test passed.")

    def test_qr_generation_structure(self):
        print("\nTesting QR Generation Structure...")
        generator = QRCodeGenerator()
        generator.set_excel_path(self.excel_path)
        generator.set_output_path("test_qr")
        
        # Mock crypto to avoid key requirement
        generator.crypto = MagicMock()
        generator.crypto.encrypt_data.return_value = b"encrypted_data"

        generator.generate_batch_qr_codes()

        expected_path = os.path.join("test_qr", "Section A", "123456789012", "123456789012.png")
        self.assertTrue(os.path.exists(expected_path), f"QR code not found at {expected_path}")
        print("QR generation structure test passed.")

if __name__ == '__main__':
    unittest.main()
