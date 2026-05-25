from flask import Flask,render_template,request,redirect,url_for,flash,abort
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager,UserMixin,login_user,login_required,logout_user,current_user
from flask_bcrypt import Bcrypt
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from flask_wtf.csrf import CSRFProtect
from flask_talisman import Talisman

import os
import datetime
from datetime import timedelta
import secrets
import re
from cryptography.fernet import Fernet
from flask import send_file
import hashlib
import time
import hmac
import magic
import requests
import io
from io import BytesIO
import cloudinary
import cloudinary.uploader
from cloudinary.utils import cloudinary_url
from dotenv import load_dotenv

load_dotenv()

# Load encryption key from environment variables
encryption_key = os.getenv("FERNET_KEY").encode()
cipher=Fernet(encryption_key)

# Secret for token generation (loaded from environment)
SECRET = os.getenv("DOWNLOAD_SECRET")

# Cloudinary configuration (reads from environment)
cloudinary.config(
    cloud_name=os.getenv('CLOUDINARY_CLOUD_NAME'),
    api_key=os.getenv('CLOUDINARY_API_KEY'),
    api_secret=os.getenv('CLOUDINARY_API_SECRET'),
    secure=True
)

# Token generator function
def generate_token(file_id):

    expiry = int(time.time()) + 300

    data = f"{file_id}:{expiry}"

    signature = hmac.new(
        SECRET.encode(),
        data.encode(),
        hashlib.sha256
    ).hexdigest()

    return f"{data}:{signature}"

app=Flask(__name__)

app.config["SECRET_KEY"]=os.getenv("SECRET_KEY")

# Secure session cookie configuration
app.config["SESSION_COOKIE_SECURE"]=False

app.config["SESSION_COOKIE_HTTPONLY"]=True

app.config["SESSION_COOKIE_SAMESITE"]="Lax"

# Use DATABASE_URL from environment for PostgreSQL in production, fall back to local SQLite for dev
default_sqlite_path = os.path.join(os.path.dirname(__file__), 'instance', 'database.db')
os.makedirs(os.path.dirname(default_sqlite_path), exist_ok=True)
app.config['SQLALCHEMY_DATABASE_URI'] = os.getenv('DATABASE_URL', f'sqlite:///{default_sqlite_path}')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db=SQLAlchemy(app)

bcrypt=Bcrypt(app)

limiter=Limiter(
    key_func=get_remote_address,
    app=app
)

csrf=CSRFProtect(app)

Talisman(
    app,
    force_https=False,
    strict_transport_security=True,
    strict_transport_security_max_age=31536000,
    content_security_policy={
        "default-src": "'self'",
        "script-src": "'self' https://cdn.jsdelivr.net",
        "style-src": "'self' https://cdn.jsdelivr.net 'unsafe-inline'",
        "font-src": "'self' https://cdn.jsdelivr.net",
        "img-src": "'self' data: https:",
        "connect-src": "'self'"
    }
)

login_manager=LoginManager()

login_manager.init_app(app)

login_manager.login_view="login"


# User model
class User(
    UserMixin,
    db.Model
):

    id=db.Column(
        db.Integer,
        primary_key=True
    )

    username=db.Column(
        db.String(100),
        unique=True
    )

    email=db.Column(
        db.String(100),
        unique=True
    )

    password=db.Column(
        db.String(200)
    )

    role=db.Column(
        db.String(20),
        default="user"
    )

    failed_attempts=db.Column(
        db.Integer,
        default=0
    )

    reset_token=db.Column(
        db.String(200),
        nullable=True
    )

#file model
class File(
    db.Model
):

    id=db.Column(
        db.Integer,
        primary_key=True
    )

    filename=db.Column(
        db.String(200)
    )

    encrypted_name=db.Column(
        db.String(200)
    )

    # Cloudinary storage details (if files stored in cloud)
    cloud_public_id = db.Column(
        db.String(255),
        nullable=True
    )

    cloud_url = db.Column(
        db.String(1024),
        nullable=True
    )

    storage_url = db.Column(
        db.String(500),
        nullable=True
    )

    cloud_resource_type = db.Column(
        db.String(20),
        default='raw'
    )

    public_id = db.Column(
        db.String(255),
        nullable=True
    )

    owner_id=db.Column(
        db.Integer,
        db.ForeignKey("user.id")
    )

    upload_time=db.Column(
        db.String(100)
    )

    file_hash=db.Column(
    db.String(300)
)


# Log model for audit logging
class Log(
    db.Model
):

    id=db.Column(
        db.Integer,
        primary_key=True
    )

    user_id=db.Column(
        db.Integer
    )

    action=db.Column(
        db.String(100)
    )

    timestamp=db.Column(
        db.String(100)
    )


# Helper function to add logs
def add_log(user_id, action):

    log = Log(
        user_id=user_id,
        action=action,
        timestamp=str(datetime.datetime.now())
    )

    db.session.add(log)
    db.session.commit()


# Notification model for simple notifications
class Notification(
    db.Model
):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    user_id = db.Column(
        db.Integer
    )

    message = db.Column(
        db.String(255)
    )

    read = db.Column(
        db.Boolean,
        default=False
    )
    created_at = db.Column(
        db.DateTime,
        default=datetime.datetime.utcnow
    )
    file_id = db.Column(
        db.Integer,
        nullable=True
    )
    share_id = db.Column(
        db.Integer,
        nullable=True
    )


# SharedFile model for file sharing
class SharedFile(
    db.Model
):

    id = db.Column(
        db.Integer,
        primary_key=True
    )

    file_id = db.Column(
        db.Integer,
        db.ForeignKey(
            'file.id'
        )
    )

    shared_by = db.Column(
        db.Integer,
        db.ForeignKey(
            'user.id'
        )
    )

    shared_with = db.Column(
        db.Integer,
        db.ForeignKey(
            'user.id'
        )
    )

    permission = db.Column(
        db.String(20)
    )

    expiry = db.Column(
        db.DateTime
    )
    
    downloads = db.Column(
        db.Integer,
        default=0
    )

    share_token = db.Column(
        db.String(255),
        nullable=True
    )


# User loader for Flask-Login
@login_manager.user_loader

def load_user(user_id):

    return User.query.get(
        int(user_id)
    )


@app.route("/")

def home():

    return redirect(
        "/login"
    )

#register route
@app.route(
    "/register",
    methods=["GET","POST"]
)

def register():

    if request.method=="POST":

        username=request.form["username"]

        email=request.form["email"]

        raw_password=request.form["password"]

        # Password policy validation
        pattern=(
            r'^(?=.*[a-z])'
            r'(?=.*[A-Z])'
            r'(?=.*\d)'
            r'(?=.*[@$!%*?&])'
            r'.{8,}$'
        )

        if not re.match(pattern, raw_password):
            flash('Password must contain uppercase, lowercase, number, special character and 8+ chars')
            return redirect('/register')

        password=bcrypt.generate_password_hash(raw_password).decode()


        user=User(
            username=username,
            email=email,
            password=password
        )


        db.session.add(user)

        db.session.commit()

        return redirect(
            "/login"
        )


    return render_template(
        "register.html"
    )


#login route
@app.route(
    "/login",
    methods=["GET","POST"]
)
@limiter.limit("5 per minute")
def login():

    if request.method=="POST":

        email=request.form["email"]

        password=request.form["password"]

        user=User.query.filter_by(
            email=email
        ).first()

        if user:

            if user.failed_attempts >= 5:
                return "Account locked due to multiple failed attempts"

            if bcrypt.check_password_hash(user.password, password):

                user.failed_attempts = 0

                db.session.commit()

                login_user(user)

                add_log(user.id, "LOGIN SUCCESS")

                return redirect(
                    "/dashboard"
                )

            else:

                user.failed_attempts += 1

                db.session.commit()

                add_log(user.id, "FAILED LOGIN")

                return "Invalid credentials"


    return render_template(
        "login.html"
    )


#dashboard route
@app.route(
    "/dashboard"
)

@login_required

def dashboard():

    return render_template(
        "dashboard.html",
        username=current_user.username
    )



#logout route
@app.route(
    "/logout"
)

@login_required

def logout():

    logout_user()

    return redirect(
        "/login"
    )

#upload route
@app.route(
    "/upload",
    methods=["POST"]
)

@login_required

def upload():

    file=request.files["file"]

    filename=file.filename

    # Real file type validation using magic bytes
    allowed_mimes = [
        "application/pdf",
        "image/png",
        "image/jpeg",
        "text/plain",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ]

    # Dangerous file signatures to block
    dangerous_signatures = [
        b'MZ',           # Windows PE executable
        b'\x7FELF',      # Linux ELF executable
        b'#!',           # Script/Shebang
        b'\xCA\xFE\xBA\xBE',  # Java class file
        b'\xFE\xED\xFA',  # Mach-O (macOS executable)
        b'PK\x03\x04',   # Zip/Jar (if not DOCX with specific structure)
    ]

    # Read file header to detect actual MIME type
    file_header=file.read(2048)
    file.seek(0)

    # Check for dangerous signatures first
    for signature in dangerous_signatures:
        if file_header.startswith(signature):
            flash("Invalid file type. Only PDF, PNG, JPG, TXT, and DOCX allowed.")
            return redirect("/upload_page")

    mime=magic.from_buffer(file_header, mime=True)

    if mime not in allowed_mimes:
        flash("Invalid file type. Only PDF, PNG, JPG, TXT, and DOCX allowed.")
        return redirect("/upload_page")

    data=file.read()

    file_hash = hashlib.sha256(data).hexdigest()

    encrypted_data=cipher.encrypt(data)

    encrypted_name=filename+".enc"

    # Upload encrypted bytes to Cloudinary as a raw resource
    public_id = f"{current_user.id}/{secrets.token_urlsafe(8)}_{encrypted_name}"
    try:
        upload_result = cloudinary.uploader.upload(
            io.BytesIO(encrypted_data),
            public_id=public_id,
            resource_type='raw',
            overwrite=True
        )
        cloud_url = upload_result.get('secure_url') or upload_result.get('url')
        cloud_public_id = upload_result.get('public_id')
        storage_url = cloud_url
    except Exception:
        flash('Upload failed')
        return redirect('/upload_page')


    new_file=File(

    filename=filename,

    encrypted_name=encrypted_name,

    cloud_public_id=cloud_public_id,

    cloud_url=cloud_url,

    storage_url=storage_url,

    public_id=cloud_public_id,

    owner_id=current_user.id,

    upload_time=str(datetime.datetime.now()),

    file_hash=file_hash

)


    db.session.add(new_file)

    db.session.commit()

    add_log(current_user.id, "UPLOAD FILE")

    return "File uploaded and encrypted successfully"

#upload route 2
@app.route("/upload_page")

@login_required

def upload_page():

    return render_template("upload.html")



#File listing route
@app.route("/files")

@login_required

def files():

    if current_user.role == "admin":

        user_files = File.query.all()

    else:

        user_files = File.query.filter_by(
            owner_id=current_user.id
        ).all()

    # count unread notifications for UI
    unread_notifications = Notification.query.filter_by(user_id=current_user.id, read=False).count()

    file_links = []

    for file in user_files:

        token = generate_token(file.id)

        # gather shares created by the current user for this file
        file_shares = []
        shares = SharedFile.query.filter_by(file_id=file.id, shared_by=current_user.id).all()
        for s in shares:
            shared_with_user = User.query.get(s.shared_with)
            file_shares.append({
                'share_id': s.id,
                'shared_with': shared_with_user.username if shared_with_user else s.shared_with,
                'permission': s.permission,
                'downloads': s.downloads,
                'expiry': s.expiry,
                'token': s.share_token
            })

        file_links.append({
            "id": file.id,
            "filename": file.filename,
            "upload_time": file.upload_time,
            "token": token,
            "shares": file_shares
        })

    return render_template(
        "files.html",
        files=file_links,
        unread_notifications=unread_notifications
    )

#File download route
@app.route("/download", methods=["POST"])

@login_required

def download():

    file_id = request.form["file_id"]

    file = File.query.get(file_id)

    if not file:
        abort(404)

    # Check if user owns the file OR has valid shared access
    is_owner = file.owner_id == current_user.id
    
    is_shared = False
    if not is_owner:
        share = SharedFile.query.filter_by(
            file_id=file.id,
            shared_with=current_user.id
        ).first()
        
        if share and share.permission == "download":
            # Check if share has expired
            if share.expiry and share.expiry < datetime.datetime.utcnow():
                flash("This shared file has expired")
                abort(403)
            is_shared = True
    
    if not is_owner and not is_shared:
        abort(403)

    token = generate_token(file_id)

    return redirect(url_for("secure_download", token=token))



#File deletion route
@app.route("/delete", methods=["POST"])

@login_required

def delete():

    file_id = request.form["file_id"]

    file = File.query.get(file_id)

    if not file or file.owner_id != current_user.id:
        abort(403)

    encrypted_path = os.path.join(
        "encrypted",
        file.encrypted_name
    )

    # Remove from Cloudinary if stored there
    try:
        if file.public_id:
            cloudinary.uploader.destroy(file.public_id, resource_type='raw')
    except Exception:
        pass

    db.session.delete(file)
    db.session.commit()

    add_log(current_user.id, "DELETE FILE")

    return redirect("/files")

#Admin route to view logs
@app.route("/logs")

@login_required

def logs():

    if current_user.role != "admin":
        return "Access denied"

    all_logs = Log.query.all()

    return render_template(
        "logs.html",
        logs=all_logs
    )

@app.route("/make_admin/<email>")

def make_admin(email):

    user = User.query.filter_by(email=email).first()

    user.role = "admin"

    db.session.commit()

    return "Admin created"

# Analytics route (STEP 7A.1)
@app.route("/analytics")

@login_required

def analytics():

    uploads = Log.query.filter_by(
        user_id=current_user.id,
        action="UPLOAD FILE"
    ).count()

    downloads = Log.query.filter_by(
        user_id=current_user.id,
        action="DOWNLOAD FILE"
    ).count()

    logins = Log.query.filter_by(
        user_id=current_user.id,
        action="LOGIN SUCCESS"
    ).count()

    return render_template(
        "analytics.html",
        uploads=uploads,
        downloads=downloads,
        logins=logins
    )

# Secure download with token (STEP 7B.3)
@app.route("/secure_download/<token>")

@login_required

def secure_download(token):

    file_id, expiry, signature = token.split(":")[:3]

    data = f"{file_id}:{expiry}"

    expected = hmac.new(
        SECRET.encode(),
        data.encode(),
        hashlib.sha256
    ).hexdigest()

    if expected != signature:
        return "Invalid link"

    if int(expiry) < int(time.time()):
        return "Link expired"

    file = File.query.get(file_id)

    if not file:
        abort(404)

    # Check ownership or valid share
    is_owner = file.owner_id == current_user.id
    
    is_shared_valid = False
    if not is_owner:
        share = SharedFile.query.filter_by(
            file_id=file.id,
            shared_with=current_user.id
        ).first()
        
        if share:
            # Enforce share expiry
            if share.expiry and share.expiry < datetime.datetime.utcnow():
                flash("This shared file has expired")
                # log expiry
                add_log(share.shared_by, f'SHARE EXPIRED {share.id}')
                abort(403)
            if share.permission == "download":
                is_shared_valid = True
    
    if not is_owner and not is_shared_valid:
        abort(403)

    # Retrieve encrypted bytes from Cloudinary (no local fallback)
    if not getattr(file, 'storage_url', None):
        abort(404)

    try:
        resp = requests.get(file.storage_url, timeout=15)
        if resp.status_code != 200:
            abort(404)
        encrypted_data = resp.content
    except Exception:
        abort(404)

    decrypted = cipher.decrypt(encrypted_data)

    current_hash = hashlib.sha256(decrypted).hexdigest()

    if current_hash != file.file_hash:
        return "File integrity compromised"

    # Stream decrypted file from memory
    mem = BytesIO(decrypted)
    mem.seek(0)

    # Increment share download counter if this was a shared download
    if not is_owner and share:
        try:
            share.downloads = (share.downloads or 0) + 1
            db.session.commit()
        except Exception:
            db.session.rollback()

    add_log(current_user.id, f'DOWNLOAD FILE {file.id}')

    return send_file(mem, as_attachment=True, download_name=file.filename)

# Password change route (STEP 7C.1)
@app.route("/change_password", methods=["POST"])

@login_required

def change_password():

    old = request.form.get("old", "")

    new = request.form.get("new", "")
    
    confirm_new = request.form.get("confirm_new", "")

    if not bcrypt.check_password_hash(current_user.password, old):
        return render_template("dashboard.html", error="Current password is incorrect")
    
    if len(new) < 6:
        return render_template("dashboard.html", error="New password must be at least 6 characters")
    
    if new != confirm_new:
        return render_template("dashboard.html", error="New passwords do not match")

    current_user.password = bcrypt.generate_password_hash(new).decode()

    db.session.commit()

    return render_template("dashboard.html", success="Password updated successfully")

# Admin statistics (STEP 7D.1)
@app.route("/admin_stats")

@login_required

def admin_stats():

    if current_user.role != "admin":
        return "Access denied"

    total_users = User.query.count()

    total_files = File.query.count()

    total_logs = Log.query.count()

    return render_template(
        "admin_stats.html",
        users=total_users,
        files=total_files,
        logs=total_logs
    )

@app.route("/admin_users")
@login_required
def admin_users():
    if current_user.role != "admin":
        return "Access denied"
    
    users = User.query.all()
    return render_template("admin_users.html", users=users)

@app.route("/admin_files")
@login_required
def admin_files():
    if current_user.role != "admin":
        return "Access denied"
    
    files = File.query.all()
    return render_template("admin_files.html", files=files)

@app.route("/admin_logs_detail")
@login_required
def admin_logs_detail():
    if current_user.role != "admin":
        return "Access denied"
    
    logs = Log.query.all()
    return render_template("admin_logs_detail.html", logs=logs)

# Forgot password route
@app.route("/forgot_password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form["email"]
        user = User.query.filter_by(email=email).first()
        
        if user:
            # Generate a secure reset token
            reset_token = secrets.token_urlsafe(32)
            user.reset_token = reset_token
            db.session.commit()
            
            return render_template(
                "reset_password_email.html",
                reset_token=reset_token,
                email=email
            )
        else:
            return render_template(
                "forgot_password.html",
                error="Email not found. Please check and try again."
            )
    
    return render_template("forgot_password.html")

# Reset password route
@app.route("/reset_password/<token>", methods=["GET", "POST"])
def reset_password(token):
    user = User.query.filter_by(reset_token=token).first()
    
    if not user:
        return render_template(
            "reset_password.html",
            error="Invalid or expired reset link"
        )
    
    if request.method == "POST":
        new_password = request.form["password"]
        confirm_password = request.form["confirm_password"]
        
        if new_password != confirm_password:
            return render_template(
                "reset_password.html",
                error="Passwords do not match"
            )
        
        if len(new_password) < 6:
            return render_template(
                "reset_password.html",
                error="Password must be at least 6 characters"
            )
        
        # Update password and clear reset token
        user.password = bcrypt.generate_password_hash(new_password).decode()
        user.reset_token = None
        db.session.commit()
        
        return render_template(
            "reset_password_success.html"
        )
    
    return render_template("reset_password.html", token=token)

# Share file route
@app.route('/share/<int:file_id>', methods=['GET', 'POST'])
@login_required
def share(file_id):
    file = db.session.get(File, file_id)
    
    if not file or file.owner_id != current_user.id:
        abort(403)
    
    if request.method == 'POST':
        username = request.form['username']
        permission = request.form['permission']
        
        user = User.query.filter_by(username=username).first()
        
        if not user:
            flash('User not found')
            return redirect(request.url)
        
        if user.id == current_user.id:
            flash('Cannot share with yourself')
            return redirect(request.url)
        
        token = secrets.token_urlsafe(32)

        # Check for existing share (idempotency)
        existing = SharedFile.query.filter_by(
            file_id=file.id,
            shared_with=user.id
        ).first()

        if existing:
            # Update existing share rather than creating duplicate
            existing.permission = permission
            existing.expiry = datetime.datetime.utcnow() + timedelta(days=7)
            if not existing.share_token:
                existing.share_token = token
            db.session.commit()

            # Create a notification for the recipient about update
            note = Notification(
                user_id=user.id,
                message=f'{current_user.username} updated sharing for file: {file.filename}',
                file_id=file.id,
                share_id=existing.id
            )
            db.session.add(note)
            db.session.commit()

            flash('Share updated successfully')
            add_log(current_user.id, f'UPDATED SHARE {existing.id} FOR FILE {file.id} WITH USER {user.id}')
            return redirect(url_for('files'))

        # Create a new share
        shared = SharedFile(
            file_id=file.id,
            shared_by=current_user.id,
            shared_with=user.id,
            permission=permission,
            expiry=datetime.datetime.utcnow() + timedelta(days=7),
            share_token=token
        )

        db.session.add(shared)
        db.session.commit()  # commit so shared.id is populated

        # Create a notification for the recipient, include file/share ids
        note = Notification(
            user_id=user.id,
            message=f'{current_user.username} shared a file: {file.filename}',
            file_id=file.id,
            share_id=shared.id
        )

        db.session.add(note)
        db.session.commit()

        flash('Shared successfully')
        add_log(current_user.id, f'SHARED FILE {file.id} WITH USER {user.id}')

        return redirect(url_for('files'))
    
    return render_template('share.html', file=file)

# View shared files route
@app.route('/shared')
@login_required
def shared():
    # count unread notifications for UI
    unread_notifications = Notification.query.filter_by(user_id=current_user.id, read=False).count()

    shared_files = SharedFile.query.filter_by(
        shared_with=current_user.id
    ).all()
    
    files_info = []
    for share in shared_files:
        file = File.query.get(share.file_id)
        owner = User.query.get(share.shared_by)
        if file:
            files_info.append({
                'id': share.id,
                'file_id': file.id,
                'filename': file.filename,
                'owner': owner.username,
                'permission': share.permission,
                'downloads': share.downloads,
                'share_token': share.share_token,
                'expiry': share.expiry,
                'upload_time': file.upload_time
            })
    
    return render_template('shared.html', files=files_info, unread_notifications=unread_notifications)


@app.route('/notifications')
@login_required
def notifications():
    notes = Notification.query.filter_by(user_id=current_user.id).order_by(Notification.id.desc()).all()
    return render_template('notifications.html', notifications=notes)


@app.route('/notifications/mark_read/<int:note_id>')
@login_required
def mark_notification_read(note_id):
    note = db.session.get(Notification, note_id)
    if not note or note.user_id != current_user.id:
        abort(404)
    note.read = True
    db.session.commit()
    return redirect(url_for('notifications'))


@app.route('/notifications_api')
@login_required
def notifications_api():
    notes = Notification.query.filter_by(user_id=current_user.id).order_by(Notification.id.desc()).limit(10).all()
    out = []
    for n in notes:
        out.append({
            'id': n.id,
            'message': n.message,
            'read': bool(n.read),
            'created_at': n.created_at.isoformat() if n.created_at else None,
            'file_id': n.file_id,
            'share_id': n.share_id
        })
    from flask import jsonify
    return jsonify(out)
    


@app.route('/revoke/<int:share_id>')
@login_required
def revoke(share_id):

    share = db.session.get(SharedFile, share_id)

    if not share:
        abort(404)

    if share.shared_by != current_user.id:
        abort(403)

    # notify recipient (optional) with file info
    file = File.query.get(share.file_id)
    note = Notification(
        user_id=share.shared_with,
        message=f'{current_user.username} revoked access to "{file.filename}"',
        file_id=file.id,
        share_id=share.id
    )
    db.session.add(note)

    db.session.delete(share)
    db.session.commit()

    flash('Access revoked')
    add_log(current_user.id, f'REVOKED SHARE {share_id}')

    return redirect('/files')

if __name__ == "__main__":

    with app.app_context():
        db.create_all()
        print("Database tables created")

    app.run(debug=True)