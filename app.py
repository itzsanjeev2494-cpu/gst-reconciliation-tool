from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import subprocess

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


# Home Page
@app.route("/")
def home():
    return render_template("index.html")


# Run Reconciliation
@app.route("/run", methods=["POST"])
def run_reconciliation():
    purchase_file = request.files["purchase_file"]
    gstr_file = request.files["gstr_file"]
    tax_month = request.form["tax_month"]

    # Save uploaded files
    purchase_path = os.path.join(UPLOAD_FOLDER, purchase_file.filename)
    gstr_path = os.path.join(UPLOAD_FOLDER, gstr_file.filename)

    purchase_file.save(purchase_path)
    gstr_file.save(gstr_path)

    # Run reconciliation script
    command = [
        "python",
        "gst_reconciliation.py",
        purchase_path,
        gstr_path,
        tax_month
    ]
old_report = os.path.join(
    "reports",
    f"GST_Reconciliation_Report_{tax_month}.xlsx"
)

if os.path.exists(old_report):
    os.remove(old_report)
    subprocess.run(command, check=True)

    # Correct output file path
  output_file = os.path.join(
    "reports",
    f"GST_Reconciliation_Report_{tax_month}.xlsx"

    )

    # Check if file exists
    if not os.path.exists(output_file):
        return f"Error: File not found -> {output_file}"

    # Download file
    return send_file(
        output_file,
        as_attachment=True
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)