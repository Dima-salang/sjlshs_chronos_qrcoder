# Generate QR Code for given excel file
import qrcode
import os
import json


def generate_qr_code(row):
    # get student id
    student_id = row['Student ID']
    # get student name
    student_name = row['Student Name']
    # get year
    year = row['Year']
    # get section
    section = row['Section']

    # get json data
    json_data = json.dumps(row.to_dict())
    print(json_data)

    # check if the qr code already exists
    if os.path.exists(f'qr/{year}/{student_id}_{year}_{section}.png'):
        print(f'QR Code already exists for {student_id}_{year}_{section}.')
        return
    else:
        os.makedirs(f'qr/{year}', exist_ok=True)
    
    qr = qrcode.QRCode(version=30, box_size=10, border=4)
    qr.add_data(json_data)
    qr.make(fit=True)
    img = qr.make_image(fill='black', back_color='white')
    img.save(f'qr/{year}/{student_id}_{year}_{section}.png')

def generate_batch_qr_code(df):
    # Generate QR Code for each row
    for index, row in df.iterrows():
        generate_qr_code(row)


    
def delete_qr_code(year, student_id, section):
    # check if the qr code exists
    if os.path.exists(f'qr/{year}/{student_id}_{year}_{section}.png'):
        os.remove(f'qr/{year}/{student_id}_{year}_{section}.png')
        print(f'QR Code deleted for {student_id}_{year}_{section}.')
    else:
        print(f'QR Code not found for {student_id}_{year}_{section}.')





