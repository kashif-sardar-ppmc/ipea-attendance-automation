import json
from flask_mail import Mail, Message
import os
import requests  # Make sure to: pip install requests
from datetime import datetime
from flask import Flask, render_template, request, redirect, url_for, flash, session
import pandas as pd
from rapidfuzz import process,fuzz
from reporting.dynamic import generate_attendance_pdf,generate_executive_pdf
emails= pd.read_csv("name and emails.csv")  # Load your CSV data into a DataFrame
print(emails.head())  # Print the first few rows to verify it's loaded correctly
app = Flask(__name__)
app.secret_key = "super_secret_key"

USER = {"username": "admin", "password": "1234"}

# --- LOGIN & LOGOUT ROUTES ---

@app.route("/", methods=["GET", "POST"])
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"]
        password = request.form["password"]
        if username == USER["username"] and password == USER["password"]:
            session["user"] = username
            return redirect(url_for("dashboard"))
        else:
            flash("Invalid username or password", "danger")
    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("user", None)
    flash("Logged out successfully", "info")
    return redirect(url_for("login"))

# --- DASHBOARD ---

@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect(url_for("login"))

    return render_template("dashboard.html", user=session["user"])


# --- OTHER ROUTES ---

import requests
from requests.auth import HTTPBasicAuth
from datetime import datetime
import calendar


@app.route("/fetch-reports", methods=["GET", "POST"])
def fetch_reports():
    if "user" not in session:
        return redirect(url_for("login"))
  

    current_month = datetime.now().strftime("%Y-%m")
    selected_month = request.form.get("month", current_month)

    # Date Range Calculation
    year, month = map(int, selected_month.split("-"))
    last_day = calendar.monthrange(year, month)[1]
    start_date = f"{selected_month}-01"
    end_date = f"{selected_month}-{last_day}"

    API_URL = "https://apps.ppmc.gov.pk/ppmc-synergy/attendance/api/daily-report/"
    auth = HTTPBasicAuth("saad.nawaz@ppmc.gov.pk", "ppmc@1234")
    
    payload = {
        "from_date": start_date,
        "to_date": end_date,
        "department": 'Integrated Planning and Economic Analysis'
    }

    grouped_reports = {}

    try:
        response = requests.post(API_URL, auth=auth, json=payload, timeout=15)
        
        if response.status_code == 200:
            raw_data = response.json().get("data", [])
            
            # Grouping the data by Employee Name
            for record in raw_data:
                emp_name = record.get("name")
                if emp_name not in grouped_reports:
                    grouped_reports[emp_name] = {
                        "id": record.get("employee_id"),
                        "dept": record.get("department"),
                        "logs": []
                    }
                grouped_reports[emp_name]["logs"].append(record)
        else:
            flash(f"API Error: {response.status_code}", "warning")
            
    except Exception as e:
        flash(f"Connection Error: {str(e)}", "danger")

    return render_template(
        "fetch_reports.html",
        reports=grouped_reports,
        selected_month=selected_month
    )
@app.route("/replies", methods=["GET", "POST"])
def replies():
    if "user" not in session:
        return redirect(url_for("login"))
    return render_template("replies.html", replies=[], selected_month="2026-02")


# Email Configuration (Example for Gmail)
app.config['MAIL_SERVER'] = 'zimbramail.nayatel.com'
app.config['MAIL_PORT'] = 587
app.config['MAIL_USE_TLS'] = True
app.config['MAIL_USERNAME'] = 'amad.atiq@ppmc.gov.pk'
app.config['MAIL_PASSWORD'] = 'Ad@06041996.' # Use App Password, not regular password
mail = Mail(app)

@app.route('/send-email', methods=['POST'])
def send_email():
    names = request.form.getlist('employee_names[]') 
    month = request.form.get('month')
    table_data_json = request.form.get('table_data')
    table_data = json.loads(table_data_json) if table_data_json else {}
    exec={}
    for i in names:
        print(i)
        exec[i] = table_data.get(i, {})
    with open("data.json", "w") as f:
        json.dump(exec, f, indent=4)
    exec_pdf=generate_executive_pdf(exec, month)
    print(f"Generated Executive PDF: {exec_pdf}")
    # for name in table_data.keys():
    #     match = process.extractOne(
    #         name,
    #         emails['Employee Name'].tolist(),
    #         scorer=fuzz.token_set_ratio
    #     )
    #     pdf=generate_attendance_pdf(table_data[name], name, month)
    #     print(f"Generated PDF for {name}: {pdf}")
    #     if match and match[1] >= 70:  # Threshold for a good match
    #         recipient_email = emails.loc[emails['Employee Name'] == match[0], 'Email'].values[0]
    #         try:
    #             msg = Message(f"Attendance Report - {month}",
    #                             sender="amad.atiq@ppmc.gov.pk",
    #                             recipients=[recipient_email])
                
    #             msg.body = f"""
    #                 Dear {name},

                    

    #                 Please find attached your monthly attendance report for {month}. You are requested to review the record and provide reasons, in the comment section, for any late arrivals, working hours less than 480 minutes (8 hours), and absences recorded during the month.

                    

    #                 Kindly submit your response within two working days to facilitate timely record finalization.
    #                 """
    #             if pdf and os.path.exists(pdf):
    #                 with open(pdf, "rb") as fp:
    #                     msg.attach(
    #                         filename=os.path.basename(pdf),
    #                         content_type="application/pdf",
    #                         data=fp.read()
    #                     )
    #             mail.send(msg)
    #               # Clean up the generated PDF after sending
    #             flash(f"Report emailed successfully to {name}!", "success")
    #         except Exception as e:
    #             flash(f"Error sending email: {str(e)}", "danger")
    #     else:
    #         print(f"No good match found for '{name}'")

    return redirect('/fetch-reports')
if __name__ == "__main__":
    app.run(debug=True, port=5000)