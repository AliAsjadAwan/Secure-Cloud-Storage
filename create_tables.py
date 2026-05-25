from app import app, db

with app.app_context():
    try:
        db.create_all()
        print('TABLES CREATED')
    except Exception as e:
        print('ERROR:', e)
