import pandas as pd
from qr_generator import generate_batch_qr_code
import os



def main():
    # Read the excel file
    df = pd.read_excel('excel/data.xlsx')
    
    # Create a directory for the QR codes
    os.makedirs('qr', exist_ok=True)

    # Generate QR Code for each row
    generate_batch_qr_code(df)

if __name__ == "__main__":
    main()