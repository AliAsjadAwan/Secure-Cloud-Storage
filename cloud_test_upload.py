import io
from app import cloudinary, cipher

# create some example data and encrypt it
plaintext = b'Test upload from CI agent'
encrypted = cipher.encrypt(plaintext)

# upload encrypted bytes to Cloudinary as raw
try:
    res = cloudinary.uploader.upload(io.BytesIO(encrypted), resource_type='raw', folder='encrypted_storage')
    print('OK', res.get('secure_url'), res.get('public_id'))
except Exception as e:
    print('ERROR', e)
