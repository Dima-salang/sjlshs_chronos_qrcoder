import qrcode
import os
import json
import base64
import io
from PIL import Image
from typing import Optional, Dict, Any, Union, Tuple, List
import pandas as pd
from qr_crypto import QRCodeCrypto

class QRCodeGenerator:
    def __init__(self, encryption_key: Optional[bytes] = None):
        """Initialize the QR code generator.
        
        Args:
            encryption_key: Optional 32-byte key for encryption. If None, no encryption is used.
        """
        self.excel_path: Optional[str] = None
        self.output_path: str = os.path.join(os.getcwd(), 'qr')
        self.crypto = QRCodeCrypto(encryption_key) if encryption_key else None
    

    def set_excel_path(self, path: str) -> None:
        """Set the path to the Excel file containing student data.
        
        Args:
            path: Path to the Excel file
        """
        if path and os.path.isfile(path):
            self.excel_path = path
    
    def set_output_path(self, path: str) -> None:
        """Set the output directory for generated QR codes.
        
        Args:
            path: Path to the output directory
        """
        if path:
            self.output_path = path
            os.makedirs(self.output_path, exist_ok=True)

    def read_excel(self) -> pd.DataFrame:
        """Read and validate the Excel file with student data.
        
        Returns:
            DataFrame containing student data
            
        Raises:
            FileNotFoundError: If Excel file is not found
            ValueError: If required columns are missing
            Exception: For other errors during Excel reading
        """
        if not self.excel_path or not os.path.isfile(self.excel_path):
            raise FileNotFoundError("Excel file not found or not specified.")
        
        try:
            df = pd.read_excel(self.excel_path)
            
            # Check for required columns
            required_columns = ['Student ID', 'Section']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                raise ValueError(f"Missing required columns in Excel file: {', '.join(missing_columns)}")
                
            # Convert Student ID to string and clean data
            df['Student ID'] = df['Student ID'].astype(str).str.strip()
            df['Section'] = df['Section'].astype(str).str.strip()
            
            # Remove any rows with missing required data
            df = df.dropna(subset=required_columns, how='any')
            
            return df
            
        except Exception as e:
            raise Exception(f"Error reading Excel file: {str(e)}")
    
    
    def create_qr_code(self, data: Dict[str, Any], output_path: str) -> bool:
        """Generate a QR code with the given data and save it to the specified path.
        
        Args:
            data: Dictionary containing the data to encode in the QR code
            output_path: Path where to save the generated QR code
            
        Returns:
            True if successful, False otherwise
        """
        try:
            # Create output directory if it doesn't exist
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            
            
            qr_data = self.crypto.encrypt_data(data)
            
            # Generate QR code with automatic version selection
            qr = qrcode.QRCode(
                version=None,
                error_correction=qrcode.constants.ERROR_CORRECT_M,
                box_size=6,
                border=2,
            )
            qr.add_data(qr_data)
            qr.make(fit=True)
            
            # Create QR code image
            qr_img = qr.make_image(fill_color="blue", back_color="white").convert('RGB')
            
            # Save the QR code
            qr_img.save(output_path, 'PNG')
            return True
            
        except Exception as e:
            print(f"Error generating QR code: {str(e)}")
            return False
    
    def generate_qr_code(self, student_data: Dict[str, Any]) -> Tuple[bool, str]:
        """Generate a QR code for a single student.
        
        Args:
            student_data: Dictionary containing student data with these required keys:
                - Student ID (str)
                - Student Name (str)
                - Year (str)
                - Section (str)
            
        Returns:
            Tuple of (success: bool, message: str)
        """
        try:
            # Extract student information
            student_id = str(student_data.get('Student ID', '')).strip()
            section = str(student_data.get('Section', '')).strip()

            # Validate required fields
            if not student_id:
                return False, f"Missing required fields for student ID: {student_id}"
            
            
            # Prepare minimal data for QR code - just the student ID
            # Other data can be looked up from the server using this ID
            qr_data = student_id
            
            
            # Create output directory path
            # Structure: Output/Section/StudentID/StudentID.png
            student_dir = os.path.join(self.output_path, section, student_id)
            os.makedirs(student_dir, exist_ok=True)
            output_path = os.path.join(student_dir, f"{student_id}.png")
            
            # Check if QR code already exists
            if os.path.exists(output_path):
                return False, f"QR code already exists for {student_id}"
            
            # Generate and save QR code
            if self.create_qr_code(qr_data, output_path):
                return True, f"Generated QR code for {student_id}"
            else:
                return False, f"Failed to generate QR code for {student_id}"
                
        except Exception as e:
            return False, f"Error processing student {student_id}: {str(e)}"
    
    def generate_batch_qr_codes(self) -> Tuple[int, int, list]:
        """Generate QR codes for all students in the Excel file.
        
        Returns:
            Tuple of (success_count, failure_count, messages)
        """
        try:
            # Read and validate Excel file
            df = self.read_excel()
            total_students = len(df)
            
            if total_students == 0:
                return 0, 0, ["No valid student records found in the Excel file."]
            
            success_count = 0
            failure_count = 0
            messages = []
            
            # Process each student
            for _, row in df.iterrows():
                success, message = self.generate_qr_code(row)
                if success:
                    success_count += 1
                else:
                    failure_count += 1
                messages.append(message)
            
            return success_count, failure_count, messages
            
        except Exception as e:
            return 0, 1, [f"Error processing batch: {str(e)}"]
        



