from flask import Flask, render_template, request, redirect, send_file, make_response
from flask_sqlalchemy import SQLAlchemy
from flask_login import LoginManager, login_user, login_required, logout_user, UserMixin
from flask_bcrypt import Bcrypt
from motor_excel import generar_cruce
import os

# ======================
# APP Permisos y Configuración Prueba test
# ======================
app = Flask(__name__)

# NOTA: En producción, cambia esto por una variable de entorno: os.environ.get('SECRET_KEY')
app.config['SECRET_KEY'] = os.environ.get('SECRET_KEY', 'Yolo09876*')
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///users.db'

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"
bcrypt = Bcrypt(app)

# ======================
# MODELO
# ======================
class User(db.Model, UserMixin):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(150), unique=True, nullable=False)
    password = db.Column(db.String(255), nullable=False)

@login_manager.user_loader
def load_user(user_id):
    # Uso de session.get en lugar de query.get (más actualizado para SQLAlchemy)
    return db.session.get(User, int(user_id))

# ======================
# RUTAS
# ======================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]

        user = User.query.filter_by(username=username).first()

        if user and bcrypt.check_password_hash(user.password, password):
            login_user(user)
            return redirect("/")

        return "Credenciales inválidas"

    return render_template("login.html")

@app.route("/test")
def test():
    return "Flask funciona"

@app.route("/logout")
def logout():
    logout_user()  # Limpia la sesión del usuario actual
    response = make_response(redirect("/"))
    response.set_cookie('session', '', expires=0) # Borra la cookie de sesión de Flask
    return response

@app.route("/generar", methods=["POST"])
@login_required
def generar():
    ingresos = request.files["ingresos"]
    egresos = request.files["egresos"]

    os.makedirs("uploads", exist_ok=True)
    os.makedirs("outputs", exist_ok=True)

    ruta_ingresos = "uploads/ingresos.xlsx"
    ruta_egresos = "uploads/egresos.xlsx"
    ruta_salida = "outputs/CRUCE_FINAL.xlsx"

    ingresos.save(ruta_ingresos)
    egresos.save(ruta_egresos)

    generar_cruce(ruta_ingresos, ruta_egresos, ruta_salida)

    return send_file(ruta_salida, as_attachment=True)

# ======================
# CREAR DB Y USUARIO (Fuera del __main__)
# ======================
# Al ponerlo aquí, se ejecutará tanto en local como en el servidor de Render
with app.app_context():
    db.create_all()
    user_existente = User.query.filter_by(username="admin").first()
    if not user_existente:
        hashed_pw = bcrypt.generate_password_hash("1234").decode('utf-8')
        user = User(username="admin", password=hashed_pw)
        db.session.add(user)
        db.session.commit()
        print("Usuario admin creado con éxito")

# ======================
# RUN (Solo para desarrollo local) Sebastian 
# ======================
if __name__ == "__main__":
    app.run(debug=True)