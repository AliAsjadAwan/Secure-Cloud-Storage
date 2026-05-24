from flask import Flask,render_template,request,redirect,url_for,flash
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager,UserMixin,login_user,login_required,logout_user,current_user
from werkzeug.security import generate_password_hash,check_password_hash


app=Flask(__name__)

app.config["SECRET_KEY"]="mysecret"

app.config["SQLALCHEMY_DATABASE_URI"]="sqlite:///database.db"

db=SQLAlchemy(app)

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

        password=generate_password_hash(
            request.form["password"]
        )


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

def login():

    if request.method=="POST":

        email=request.form["email"]

        password=request.form["password"]

        user=User.query.filter_by(
            email=email
        ).first()


        if user and check_password_hash(
            user.password,
            password
        ):

            login_user(user)

            return redirect(
                "/dashboard"
            )


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


#create database tables
with app.app_context():

    db.create_all()


if __name__=="__main__":

    app.run(
        debug=True
    )