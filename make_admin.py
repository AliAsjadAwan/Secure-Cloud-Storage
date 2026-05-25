#!/usr/bin/env python
import sys
sys.path.insert(0, 'e:\\CUI\\IS\\Lab\\secure_cloud_storage')
from app import app, db, User

with app.app_context():
    user = User.query.filter_by(email='test@example.com').first()
    if user:
        user.role = 'admin'
        db.session.commit()
        print(f'User {user.username} is now admin')
    else:
        print('User not found')
