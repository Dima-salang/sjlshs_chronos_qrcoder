# Generate QR Code for given excel file
import pandas as pd
import qrcode
import os
import json

# Create a directory for the QR codes
os.makedirs('qr', exist_ok=True)

# Read the excel file
df = pd.read_excel('excel/data.xlsx')


# Generate QR Code for each row
for index, row in df.iterrows():
    # get student id
    student_id = row['Student ID']
    # get student name
    student_name = row['Student Name']
    # get year
    year = row['Year']
    # get section
    section = row['Section']

    # check if the qr code already exists
    if os.path.exists(f'qr/{student_id}_{year}_{section}.png'):
        continue

    qr = qrcode.QRCode(version=30, box_size=10, border=4)
    qr.add_data(json.dumps(row.to_dict()))
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    img.save(f'qr/{student_id}_{year}_{section}.png')




