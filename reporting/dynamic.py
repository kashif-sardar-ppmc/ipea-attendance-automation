import datetime
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import os
import tempfile
import re
import subprocess
import json

import requests
from requests.auth import HTTPBasicAuth

def fetch_attendance_from_api(from_date, to_date, department):
    url = "https://apps.ppmc.gov.pk/ppmc-synergy/attendance/api/daily-report/"

    payload = {
        "from_date": from_date,
        "to_date": to_date,
        "department": department
    }
    
    try:
        response = requests.post(
            url,
            json=payload,
            auth=HTTPBasicAuth(
                "saad.nawaz@ppmc.gov.pk",
                "ppmc@1234"
            ),
            timeout=30
        )

        response.raise_for_status()  # Raise error if not 200

        data = response.json()

        print("✅ API Data Fetched Successfully")
        return data

    except requests.exceptions.RequestException as e:
        print(f"❌ API Request Failed: {e}")
        return {}


def generate_attendance_pdf(attendance_data, month_year):
    from datetime import datetime
    import os
    import pandas as pd

    # --------- DATE CONFIG ---------
    RAMZAN_START = datetime.strptime("2026-02-19", "%Y-%m-%d").date()
    RAMZAN_END = datetime.strptime("2026-03-20", "%Y-%m-%d").date()
    FRIDAY_ACTIVE_FROM = datetime.strptime("2026-03-12", "%Y-%m-%d").date()

    # --------- POLICY FUNCTION ---------
    def get_work_policy(date):
        d = date.date()
        weekday = date.weekday()  # Monday=0 ... Friday=4
        if d >= datetime.strptime("2026-03-24", "%Y-%m-%d").date() and weekday == 4:  # Friday always off after Ramzan
            return {
            "office_start": datetime.strptime("08:20", "%H:%M").time(),
            "required_hours": 8,
            "is_off": True}
        if d >= datetime.strptime("2026-03-24", "%Y-%m-%d").date() and weekday == 4:  # Friday always off after Ramzan
            return {
            "office_start": datetime.strptime("08:20", "%H:%M").time(),
            "required_hours": 8,
            "is_off": True}
        # RAMZAN
        if RAMZAN_START <= d <= RAMZAN_END:
            if weekday == 4:  # Friday
                if d < FRIDAY_ACTIVE_FROM:
                    return {"is_off": True}
                return {
                    "office_start": datetime.strptime("09:20", "%H:%M").time(),
                    "required_hours": 3.5,
                    "is_off": False
                }
            return {
                "office_start": datetime.strptime("09:20", "%H:%M").time(),
                "required_hours": 6,
                "is_off": False
            }

        # NORMAL DAYS
        if weekday == 4 and d > RAMZAN_END:  # Friday always off
            return {"is_off": True}

        return {
            "office_start": datetime.strptime("08:50", "%H:%M").time(),
            "required_hours": 8,
            "is_off": False
        }

    # --------- OUTPUT SETUP ---------
    output_dir = 'temp'
    os.makedirs(output_dir, exist_ok=True)
    filename = f"Attendance_{month_year.replace(' ', '_')}.pdf"
    output_pdf = os.path.join(output_dir, filename)

    all_summary_data = []

    for name, records in attendance_data.items():
        df_emp = pd.DataFrame(records)
        if df_emp.empty:
            continue

        df_emp["status"] = df_emp["status"].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_emp["date"] = pd.to_datetime(df_emp["date"], errors="coerce")

        # --------- FILTER WORKING DAYS USING POLICY ---------
        working_days = df_emp[df_emp["date"].apply(lambda d: not get_work_policy(d).get("is_off", False))]
        total_days = len(working_days)

        presents = working_days["status"].str.contains("Present", case=False, na=False).sum()
        attendance_pct = round((presents / total_days * 100), 2) if total_days > 0 else 0

        late_count = 0
        total_late_minutes = 0
        total_less_minutes = 0
        valid_work_days = 0

        for _, row in working_days.iterrows():
            ci, co = row.get("check_in"), row.get("check_out")
            if not ci or not co or ci in ["-", "Absent"] or co in ["-", "Absent"]:
                continue

            try:
                ci_time = datetime.strptime(ci.split()[0], "%H:%M:%S" if ci.count(':') == 2 else "%H:%M")
                co_time = datetime.strptime(co.split()[0], "%H:%M:%S" if co.count(':') == 2 else "%H:%M")

                policy = get_work_policy(row["date"])
                if policy.get("is_off"):
                    continue

                OFFICE_START = policy["office_start"]
                REQUIRED_HOURS = policy["required_hours"]

                # --------- LATE ---------
                if ci_time.time() > OFFICE_START:
                    late_count += 1
                    late_minutes = (
                        datetime.combine(datetime.today(), ci_time.time()) -
                        datetime.combine(datetime.today(), OFFICE_START)
                    ).total_seconds() / 60
                    total_late_minutes += late_minutes

                # --------- WORK HOURS ---------
                work_duration_hours = (co_time - ci_time).total_seconds() / 3600

                # --------- SHORT HOURS ---------
                if work_duration_hours < REQUIRED_HOURS:
                    total_less_minutes += (REQUIRED_HOURS - work_duration_hours) * 60

                valid_work_days += 1

            except:
                continue

        avg_late_minutes = round(total_late_minutes / late_count, 0) if late_count > 0 else 0
        avg_less_minutes = round(total_less_minutes / valid_work_days, 0) if valid_work_days > 0 else 0

        all_summary_data.append({
            "Employee Name": name,
            "Late Count (Days)": late_count,
            "Average Late Minutes": int(avg_late_minutes),
            "Average Minutes Less Than Required Hours": int(avg_less_minutes),
            "Attendance %": attendance_pct
        })

    summary_df = pd.DataFrame(all_summary_data)
    dept_attendance = round(summary_df["Attendance %"].mean(), 2) if not summary_df.empty else 0

    # --------- HTML TABLE ---------
    table_rows = ""
    for _, row in summary_df.iterrows():
        table_rows += f"""
        <tr style="border-bottom:1px solid #ddd; text-align:center;">
            <td style="padding:8px; text-align:left;">{row['Employee Name']}</td>
            <td style="padding:8px;">{row['Late Count (Days)']}</td>
            <td style="padding:8px;">{row['Average Late Minutes']}</td>
            <td style="padding:8px;">{row['Average Minutes Less Than Required Hours']}</td>
            <td style="padding:8px;">{row['Attendance %']:.2f}%</td>
        </tr>
        """

    table_rows += f"""
    <tr style="font-weight:bold; background:#f2f2f2; text-align:center;">
        <td style="padding:10px;">Departmental Percentage</td>
        <td colspan="3"></td>
        <td>{dept_attendance:.2f}%</td>
    </tr>
    """

    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            h1 {{ color:#1f6f3e; }}
            table {{ width:100%; border-collapse:collapse; }}
            th {{
                background-color:#1f6f3e;
                color:white;
                padding:12px;
                text-align:center;
            }}
        </style>
    </head>
    <body>
        <h1>Attendance Summary</h1>
        <p><strong>Report Period:</strong> {month_year}</p>
        <table>
            <thead>
                <tr>
                    <th>Employee Name</th>
                    <th>Late Count (Days)</th>
                    <th>Average Late Minutes</th>
                    <th>Average Minutes Less Than Required Hours</th>
                    <th>Attendance %</th>
                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
    </body>
    </html>
    """

    # --------- PDF EXPORT ---------
    path_to_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'
    try:
        import pdfkit
        config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)
        pdfkit.from_string(html_content, output_pdf, configuration=config, options={'encoding': "UTF-8", 'quiet': ''})
        print(f"✅ Success! File created at: {output_pdf}")
    except Exception as e:
        print(f"❌ PDF Failed: {e}")
        with open(output_pdf.replace(".pdf", ".html"), "w", encoding="utf-8") as f:
            f.write(html_content)

    return output_pdf


def generate_executive_pdf(attendance_data, month_year):
    from datetime import datetime
    import os
    import pandas as pd

    # --------- DATE CONFIG ---------
    RAMZAN_START = datetime.strptime("2026-02-19", "%Y-%m-%d").date()
    RAMZAN_END = datetime.strptime("2026-03-20", "%Y-%m-%d").date()
    FRIDAY_ACTIVE_FROM = datetime.strptime("2026-03-12", "%Y-%m-%d").date()

    # --------- POLICY FUNCTION ---------
    def get_work_policy(date):
        d = date.date()
        weekday = date.weekday()  # Monday=0 ... Friday=4

        # RAMZAN
        if RAMZAN_START <= d <= RAMZAN_END:
            if weekday == 4:  # Friday
                if d < FRIDAY_ACTIVE_FROM:
                    return {"is_off": True}
                return {
                    "office_start": datetime.strptime("09:20", "%H:%M").time(),
                    "required_hours": 3.5,
                    "is_off": False
                }
            return {
                "office_start": datetime.strptime("09:20", "%H:%M").time(),
                "required_hours": 6,
                "is_off": False
            }

        # NORMAL DAYS
        if weekday == 4 and d>RAMZAN_END:  # Friday always off
            return {"is_off": True}

        return {
            "office_start": datetime.strptime("08:50", "%H:%M").time(),
            "required_hours": 8,
            "is_off": False
        }

    # --------- OUTPUT ---------
    output_dir = 'temp'
    os.makedirs(output_dir, exist_ok=True)
    filename = f"Executive_Summary_{month_year.replace(' ', '_')}.pdf"
    output_pdf = os.path.join(output_dir, filename)

    all_summary_data = []

    for name, records in attendance_data.items():
        df_emp = pd.DataFrame(records)
        if df_emp.empty:
            continue

        df_emp["status"] = df_emp["status"].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
        df_emp["date"] = pd.to_datetime(df_emp["date"], errors="coerce")

        # --------- FILTER WORKING DAYS USING POLICY ---------
        working_days = df_emp[
            df_emp["date"].apply(lambda d: not get_work_policy(d).get("is_off", False))
        ]

        total_days = len(working_days)

        presents = working_days["status"].str.contains("Present", case=False, na=False).sum()
        
        attendance_pct = round((presents / total_days * 100), 2) if total_days > 0 else 0

        late_count = 0
        total_late_minutes = 0
        total_less_minutes = 0
        valid_work_days = 0

        for _, row in working_days.iterrows():
            ci, co = row.get("check_in"), row.get("check_out")

            if not ci or not co or ci in ["-", "Absent"] or co in ["-", "Absent"]:
                continue

            try:
                ci_time = datetime.strptime(ci.split()[0], "%H:%M:%S" if ci.count(':') == 2 else "%H:%M")
                co_time = datetime.strptime(co.split()[0], "%H:%M:%S" if co.count(':') == 2 else "%H:%M")

                policy = get_work_policy(row["date"])

                # Skip if off (extra safety)
                if policy.get("is_off"):
                    continue

                OFFICE_START = policy["office_start"]
                REQUIRED_HOURS = policy["required_hours"]

                # --------- LATE ---------
                if ci_time.time() > OFFICE_START:
                    late_count += 1
                    late_minutes = (
                        datetime.combine(datetime.today(), ci_time.time()) -
                        datetime.combine(datetime.today(), OFFICE_START)
                    ).total_seconds() / 60
                    total_late_minutes += late_minutes

                # --------- WORK HOURS ---------
                work_duration_hours = (co_time - ci_time).total_seconds() / 3600

                # --------- SHORT HOURS ---------
                if work_duration_hours < REQUIRED_HOURS:
                    total_less_minutes += (REQUIRED_HOURS - work_duration_hours) * 60

                valid_work_days += 1

            except:
                continue

        avg_late_minutes = round(total_late_minutes / late_count, 0) if late_count > 0 else 0
        avg_less_minutes = round(total_less_minutes / valid_work_days, 0) if valid_work_days > 0 else 0

        all_summary_data.append({
            "Employee Name": name,
            "Late Count (Days)": late_count,
            "Average Late Minutes": int(avg_late_minutes),
            "Average Minutes Less Than Required Hours": int(avg_less_minutes),
            "Attendance %": attendance_pct,
            "Present":presents,
            "Total":total_days

        })

    summary_df = pd.DataFrame(all_summary_data)
    dept_attendance = round(summary_df["Attendance %"].mean(), 2) if not summary_df.empty else 0

    # --------- HTML ---------
    table_rows = ""
    for _, row in summary_df.iterrows():
        table_rows += f"""
        <tr style="border-bottom:1px solid #ddd; text-align:center;">
            <td style="padding:8px; text-align:left;">{row['Employee Name']}</td>
            <td style="padding:8px;">{row['Late Count (Days)']}</td>
            <td style="padding:8px;">{row['Average Late Minutes']}</td>
            <td style="padding:8px;">{row['Average Minutes Less Than Required Hours']}</td>
            <td style="padding:8px;">{row['Attendance %']:.2f}%</td>
            <td style="padding:8px;">{row['Present']:.2f}</td>
            <td style="padding:8px;">{row['Total']:.2f}</td>
        </tr>
        """

    table_rows += f"""
    <tr style="font-weight:bold; background:#f2f2f2; text-align:center;">
        <td style="padding:10px;">Departmental Percentage</td>
        <td colspan="3"></td>
        <td>{dept_attendance:.2f}%</td>
    </tr>
    """

    html_content = f"""
    <html>
    <head>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; }}
            h1 {{ color:#1f6f3e; }}
            table {{ width:100%; border-collapse:collapse; }}
            th {{
                background-color:#1f6f3e;
                color:white;
                padding:12px;
                text-align:center;
            }}
        </style>
    </head>
    <body>
        <h1>Executive Attendance Summary</h1>
        <p><strong>Report Period:</strong> {month_year}</p>
        <table>
            <thead>
                <tr>
                    <th>Employee Name</th>
                    <th>Late Count (Days)</th>
                    <th>Average Late Minutes</th>
                    <th>Average Minutes Less Than Required Hours</th>
                    <th>Attendance %</th>
                    <th>Presents</th>
                    <th>Total</th>

                </tr>
            </thead>
            <tbody>
                {table_rows}
            </tbody>
        </table>
    </body>
    </html>
    """

    # --------- PDF ---------
    path_to_wkhtmltopdf = r'C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe'

    try:
        import pdfkit
        config = pdfkit.configuration(wkhtmltopdf=path_to_wkhtmltopdf)
        pdfkit.from_string(html_content, output_pdf, configuration=config, options={'encoding': "UTF-8", 'quiet': ''})
        print(f"✅ Success! File created at: {output_pdf}")

    except Exception as e:
        print(f"❌ PDF Failed: {e}")
        with open(output_pdf.replace(".pdf", ".html"), "w", encoding="utf-8") as f:
            f.write(html_content)

    return output_pdf

def main():
    # Example usage
    attendance_data = fetch_attendance_from_api(
        from_date="2026-03-01",
        to_date="2026-03-25",
        department="IT"
    )
    # Generate individual report for an employee
    emp_report = generate_attendance_pdf(attendance_data["John Doe"], "John Doe", "September 2024")
    print(f"Generated Employee Report: {emp_report}")

    # Generate executive summary for all employees
    exec_report = generate_executive_pdf(attendance_data, "September 2024")
    print(f"Generated Executive Summary: {exec_report}")