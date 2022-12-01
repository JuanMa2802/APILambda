import os 
import subprocess
from flask import Flask, request, abort, render_template
from subprocess import Popen

app = Flask(__name__, template_folder="templates")

#Ruta base
@app.route("/")
def home():
    return render_template("index.html")


#Formulario de inscripción
@app.route("/api/v1/FormularioInscripcion/<string:Id>")
def FormularioInscripcion(Id):
    respuesta = subprocess.run(["python", r"C:\Users\Juan Manuel Gaviria\Desktop\APILambda\scripts\FormInscription.py", Id], capture_output=True, timeout=120)
    return "respuesta"


#Ordenes de compra
@app.route("/api/v1/OrdenesComprar/<string:Id>")
def OrdenesComprar(Id):
    respuesta = subprocess.run(["python", r"/scripts/RequestsArconsa/manage.py", "Arconsa", Id], capture_output=True)

    return respuesta.stdout


if __name__ == "__main__":
    app.run(
        debug=True, 
        host="0.0.0.0", 
        port=int(os.environ.get("PORT", 4000))
    )