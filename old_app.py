from flask import Flask, render_template, request, send_file
from cryptography.fernet import Fernet
import os

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
ENCRYPTED_FOLDER = "encrypted"
DECRYPTED_FOLDER = "decrypted"

with open("secret.key", "rb") as file:
    key = file.read()

cipher = Fernet(key)


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():

    file = request.files["file"]

    filepath = os.path.join(
        UPLOAD_FOLDER,
        file.filename
    )

    file.save(filepath)

    with open(filepath,"rb") as f:
        data = f.read()

    encrypted = cipher.encrypt(data)

    encrypted_path = os.path.join(
        ENCRYPTED_FOLDER,
        file.filename + ".enc"
    )

    with open(encrypted_path,"wb") as f:
        f.write(encrypted)

    return "Encrypted and stored securely"


@app.route("/decrypt", methods=["POST"])
def decrypt():

    filename=request.form["filename"]

    encrypted_path=os.path.join(
        ENCRYPTED_FOLDER,
        filename
    )

    with open(encrypted_path,"rb") as f:
        encrypted=f.read()

    decrypted=cipher.decrypt(encrypted)

    output=filename.replace(".enc","")

    path=os.path.join(
        DECRYPTED_FOLDER,
        output
    )

    with open(path,"wb") as f:
        f.write(decrypted)

    return send_file(path, as_attachment=True)


if __name__=="__main__":
    app.run(debug=True)