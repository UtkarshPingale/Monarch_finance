import hashlib

password = "admin123"  # Change to the password you want

password_hash = hashlib.sha256(password.encode("utf-8")).hexdigest()

print("Password:", password)
print("SHA-256:", password_hash)