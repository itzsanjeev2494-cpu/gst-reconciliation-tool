from flask import Flask, render_template, request, send_file
import os
import subprocess

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
REPORT_FOLDER = "reports"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(REPORT_FOLDER, exist_ok=True)

@app.route("/")
def home():
    return render_template("index.html")


@app.route("/run", methods=["POST"])
def run_reconciliation():
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