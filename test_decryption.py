import os
import base64
from qr_crypto import QRCodeCrypto
from qr_generator import QRCodeGenerator
from PIL import Image
import qrcode
from pyzbar.pyzbar import decode

def test_qr_encryption_decryption():
    # 1. Set up test data
    test_data = {
        'student_id': '12345',
        'name': 'Test Student',
        'year': '2023',
        'section': 'A',
        'test_field': 'This is a test message'
    }
    
    # 2. Create a test key or use existing one
    key_file = 'test_key.key'
    if os.path.exists(key_file):
        with open(key_file, 'rb') as f:
            key = base64.b64decode(f.read())
    else:
        # Generate a new key for testing
        key = os.urandom(32)  # 256-bit key
        with open(key_file, 'wb') as f:
            f.write(base64.b64encode(key))
    
    print(f"Using key: {base64.b64encode(key).decode('utf-8')}")
    
    # 3. Initialize crypto and generator
    crypto = QRCodeCrypto(key)
    qr_gen = QRCodeGenerator(key)
    
    # 4. Encrypt the test data
    encrypted_data = crypto.encrypt_data(test_data)
    print(f"\nEncrypted data (first 50 chars): {encrypted_data[:50]}...")
    
    # 5. Create a QR code with the encrypted data
    qr = qrcode.QRCode(version=1, box_size=10, border=5)
    qr.add_data(encrypted_data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white")
    
    # Save the QR code temporarily
    qr_path = 'test_qr.png'
    qr_img.save(qr_path)
    print(f"\nTest QR code saved to: {os.path.abspath(qr_path)}")
    
    # 6. Read the QR code back
    decoded_objects = decode(Image.open(qr_path))
    if not decoded_objects:
        print("Error: Could not read QR code")
        return
        
    read_data = decoded_objects[0].data.decode('utf-8')
    print(f"\nRead data from QR (first 50 chars): {read_data[:50]}...")
    
    # 7. Decrypt the data
    try:
        decrypted_data = crypto.decrypt_data(read_data)
        print("\nDecryption successful!")
        print("Original data:", test_data)
        print("Decrypted data:", decrypted_data)
        
        # 8. Verify the data matches
        if decrypted_data == test_data:
            print("\n✅ Test passed: Decrypted data matches original!")
        else:
            print("\n❌ Test failed: Decrypted data does not match original!")
            
    except Exception as e:
        print(f"\n❌ Decryption failed: {str(e)}")
    
    # Clean up
    if os.path.exists(qr_path):
        os.remove(qr_path)

if __name__ == "__main__":
    test_qr_encryption_decryption()
