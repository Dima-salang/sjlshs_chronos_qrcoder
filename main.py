import pandas as pd
from qr_generator import generate_batch_qr_code, generate_qr_code, delete_qr_code
import os

"""
If you would like to generate only a single qr code, include this in the main function
df.DataFrame()
"""


def main():
    # Read the excel file
    df = pd.read_excel('excel/data.xlsx')
    
    # Create a directory for the QR codes
    os.makedirs('qr', exist_ok=True)

    # if generating individual qr codes, use this
    #qr_data = pd.Series(
    #    {
    #       "Student ID": "1234567893",
    #        "Student Name": 'John Doe',
    #        "Year": 2025,
    #        "Section": 'A'
    #    }
    #)

    # test generate single qr code
    # generate_qr_code(qr_data)
    
    # delete qr code
    # delete_qr_code('2024', '1234567890', 'A')
    # Generate QR Code for each row
    generate_batch_qr_code(df)

if __name__ == "__main__":
    main()