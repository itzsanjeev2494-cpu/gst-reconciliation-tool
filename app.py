from flask import Flask, render_template, request, send_file, redirect, session
import os
import subprocess

app = Flask(__name__)
app.secret_key = "mygstsecretkey123"

UPLOAD_FOLDER = "uploads"
REPORT_FOLDER = "reports"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)


# LOGIN PAGE
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        password = request.form["password"]

        if password == "GST2026":
            session["logged_in"] = True
            return redirect("/")
        else:
            return "Wrong Password"

    return '''
    <form method="POST">
        <h2>GST Tool Login</h2>
        <input type="password" name="password" placeholder="Enter Password">
        <button type="submit">Login</button>
    </form>
    '''


# LOGOUT PAGE
@app.route("/logout")
def logout():
    session.pop("logged_in", None)
    return redirect("/login")


# HOME PAGE (Protected)
@app.route("/")
def home():
    if not session.get("logged_in"):
        return redirect("/login")

    return render_template("index.html")


# RUN RECONCILIATION
@app.route("/run", methods=["POST"])
def run_reconciliation():
    if not session.get("logged_in"):
        return redirect("/login")

    purchase_file = request.files["purchase_file"]
    gstr_file = request.files["gstr_file"]
    tax_month = request.form["tax_month"]

    purchase_path = os.path.join(UPLOAD_FOLDER, purchase_file.filename)
    gstr_path = os.path.join(UPLOAD_FOLDER, gstr_file.filename)

    purchase_file.save(purchase_path)
    gstr_file.save(gstr_path)

    command = [
        "python",
        "gst_reconciliation.py",
        purchase_path,
        gstr_path,
        tax_month
    ]

    subprocess.run(command, check=True)

    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

    output_file = os.path.join(
        BASE_DIR,
        f"{tax_month}_Recon",
        f"GST_Reconciliation_Report_{tax_month}.xlsx"
    )

    return send_file(output_file, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)