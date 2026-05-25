from app import app
from app import db

with app.app_context():
    try:
        conn = db.engine.connect()
        print("CONNECTED")
        conn.close()
    except Exception as e:
        print(e)
