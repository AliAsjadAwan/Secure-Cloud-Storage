from app import app, db

with app.app_context():
    db.create_all()
    print("Database created successfully!")
    print("Tables created: user, file, log, shared_file")
