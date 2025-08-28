"""
Symmetric encryption for QR codes using AES-256-GCM.
Provides fast encryption and decryption with a shared secret key.
"""
import os
import json
from base64 import b64encode, b64decode
from Crypto.Cipher import AES
from Crypto.Random import get_random_bytes
from typing import Dict, Any, Tuple, Optional

class QRCodeCrypto:
    """Handles encryption and decryption of QR code data using a shared secret key."""
    
    def __init__(self, key: Optional[bytes] = None, key_file: str = 'encryption_key.key'):
        """Initialize with an optional 32-byte key or load from file.
        
        Args:
            key: 32-byte key for AES-256. If None, will try to load from key_file.
            key_file: Path to the key file to load/save the key.
        """
        self.key_file = key_file
        
        if key is not None:
            if len(key) != 32:
                raise ValueError("Key must be 32 bytes (256 bits) for AES-256")
            self.key = key
        elif os.path.exists(key_file):
            self.key = self.load_key(key_file)
        else:
            self.key = get_random_bytes(32)  # 256-bit key for AES-256
            # Save the generated key to file
            with open(key_file, 'wb') as f:
                f.write(b64encode(self.key))
            
    def key_exists(self) -> bool:
        """Check if a key file exists."""
        return os.path.exists(self.key_file)
        
    @classmethod
    def load_key(cls, key_file: str) -> bytes:
        """Load the key from the key file.
        
        Args:
            key_file: Path to the key file
            
        Returns:
            The loaded key as bytes
        """
        with open(key_file, 'rb') as f:
            return b64decode(f.read())
            
    def get_key_base64(self) -> str:
        """Get the key as a base64 string for sharing with Flutter."""
        return b64encode(self.key).decode('utf-8')
    
    def encrypt_data(self, data: Dict[str, Any]) -> str:
        """Encrypt the data dictionary into a string.
        
        Args:
            data: Dictionary containing the data to encrypt
            
        Returns:
            Base64-encoded string containing the encrypted data and nonce
        """
        # Convert data to JSON string
        json_data = json.dumps(data, ensure_ascii=False).encode('utf-8')
        
        # Generate a random nonce
        nonce = get_random_bytes(12)  # 96 bits for GCM
        
        # Create cipher and encrypt
        cipher = AES.new(self.key, AES.MODE_GCM, nonce=nonce)
        ciphertext, tag = cipher.encrypt_and_digest(json_data)
        
        # Combine nonce, tag, and ciphertext
        encrypted_data = nonce + tag + ciphertext
        
        # Return as base64 string for easy QR code encoding
        return b64encode(encrypted_data).decode('utf-8')
    
    def decrypt_data(self, encrypted_data_str: str) -> Dict[str, Any]:
        """Decrypt the data from an encrypted string.
        
        Args:
            encrypted_data_str: Base64-encoded string from the QR code
            
        Returns:
            Decrypted data as a dictionary
            
        Raises:
            ValueError: If decryption or verification fails
        """
        try:
            # Decode the base64 string
            encrypted_data = b64decode(encrypted_data_str)
            
            # Extract nonce (first 12 bytes), tag (next 16 bytes), and ciphertext (the rest)
            nonce = encrypted_data[:12]
            tag = encrypted_data[12:28]
            ciphertext = encrypted_data[28:]
            
            # Create cipher and decrypt
            cipher = AES.new(self.key, AES.MODE_GCM, nonce=nonce)
            json_data = cipher.decrypt_and_verify(ciphertext, tag)
            
            # Convert back to dictionary
            return json.loads(json_data.decode('utf-8'))
            
        except Exception as e:
            raise ValueError("Decryption failed. The QR code may be corrupted or the key is incorrect.") from e
    
    def save_key(self, file_path: str) -> None:
        """Save the key to a file.
        
        Args:
            file_path: Path to save the key file
        """
        with open(file_path, 'wb') as f:
            f.write(self.key)
    
    @classmethod
    def load_from_file(cls, file_path: str) -> 'QRCodeCrypto':
        """Load a key from a file and return a new QRCodeCrypto instance.
        
        Args:
            file_path: Path to the key file
            
        Returns:
            New QRCodeCrypto instance with the loaded key
        """
        with open(file_path, 'rb') as f:
            key = b64decode(f.read())
        return cls(key)
    
    @classmethod
    def generate_and_save_key(cls, file_path: str) -> 'QRCodeCrypto':
        """Generate a new key, save it to a file, and return a new instance.
        
        Args:
            file_path: Path to save the new key file
            
        Returns:
            New QRCodeCrypto instance with the generated key
        """
        crypto = cls()
        crypto.save_key(file_path)
        return crypto
