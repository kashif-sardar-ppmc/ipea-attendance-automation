import pandas as pd
import matplotlib.pyplot as plt
import os


def generate_attendance_report(
        excel_file,
        template_file,
        ppmc_logo_path,
        state_logo_path,
        sheet_name,
        output_folder):

    print(f"Processing sheet: {sheet_name}")

    safe_sheet_name = sheet_name.replace(" ", "_")

    final_html = os.path.join(output_folder, f"{safe_sheet_name}_report.html")
    final_pdf = os.path.join(output_folder, f"{safe_sheet_name}_Attendance_Report.pdf")

    chart1_path = os.path.join(output_folder, f"{safe_sheet_name}_chart1.png")
    chart2_path = os.path.join(output_folder, f"{safe_sheet_name}_chart2.png")

    # -----------------------------
    # LOAD DATA
    # -----------------------------

    df = pd.read_excel(excel_file, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()

    if df.empty:
        print(f"Skipping {sheet_name} (empty sheet)")
        return

    required_columns = ["Name", "Date", "Worked Minutes", "Late Minutes", "Status"]

    if not all(col in df.columns for col in required_columns):
        print(f"Skipping {sheet_name} (missing required columns)")
        return

    # -----------------------------
    # DATE PROCESSING
    # -----------------------------

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    valid_dates = df["Date"].dropna()

    month_year = valid_dates.iloc[0].strftime("%B %Y") if not valid_dates.empty else "N/A"

    df["Date"] = df["Date"].dt.strftime("%d-%m-%Y")

    employee_name = df["Name"].iloc[0]

    # -----------------------------
    # NUMERIC CLEANING
    # -----------------------------

    numeric_cols = ["Worked Minutes", "Late Minutes", "Less Than 8 Hours"]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").round(0)

    df = df.fillna("")

    # -----------------------------
    # TABLE GENERATION
    # -----------------------------

    display_columns = [
        col for col in [
            "Date",
            "Time In",
            "Time Out",
            "Worked Minutes",
            "Late Minutes",
            "Less Than 8 Hours",
            "Status",
            "Comments"
        ] if col in df.columns
    ]

    rows_html = ""

    for index, row in df.iterrows():

        less_than_8 = pd.to_numeric(row.get("Less Than 8 Hours", ""), errors="coerce")

        if pd.notna(less_than_8) and less_than_8 > 0:
            bg_color = "#ffdddd"
        else:
            bg_color = "#f2f9f4" if index % 2 != 0 else "white"

        rows_html += f"<tr style='background:{bg_color};'>"

        for col in display_columns:
            rows_html += f"<td style='padding:5px; border:1px solid #ccc; text-align:center; font-size:12px;'>{row[col]}</td>"

        rows_html += "</tr>"

    print("Rows generated:", len(rows_html))

    # -----------------------------
    # SUMMARY
    # -----------------------------

    worked_numeric = pd.to_numeric(df["Worked Minutes"], errors="coerce")
    late_numeric = pd.to_numeric(df["Late Minutes"], errors="coerce")

    total_days = df["Date"].nunique()
    present_days = worked_numeric.gt(0).sum()
    total_worked = worked_numeric.sum()
    total_late = late_numeric.sum()
    late_days = df["Status"].astype(str).str.strip().str.lower().eq("late").sum()

    attendance_percentage = round(
        (present_days / total_days) * 100, 2) if total_days > 0 else 0

    summary_html = f"""
    <ul>
        <li><strong>Total Working Days:</strong> {total_days}</li>
        <li><strong>Days Present:</strong> {present_days}</li>
        <li><strong>Total Late Days:</strong> {late_days}</li>
        <li><strong>Total Worked Minutes:</strong> {int(total_worked)}</li>
        <li><strong>Total Late Minutes:</strong> {int(total_late)}</li>
        <li><strong>Attendance Percentage:</strong> {attendance_percentage}%</li>
    </ul>
    """

    # -----------------------------
    # CREATE CHARTS
    # -----------------------------

    plt.figure(figsize=(6, 4))
    plt.plot(df["Date"], worked_numeric)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(chart1_path)
    plt.close()

    plt.figure(figsize=(6, 4))
    plt.bar(df["Date"], late_numeric)
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig(chart2_path)
    plt.close()

    # -----------------------------
    # CONVERT PATHS TO file:///
    # -----------------------------

    def to_file_url(path):
        return "file:///" + path.replace("\\", "/")

    chart1_url = to_file_url(chart1_path)
    chart2_url = to_file_url(chart2_path)
    logo1_url = to_file_url(ppmc_logo_path)
    logo2_url = to_file_url(state_logo_path)

    # -----------------------------
    # LOAD TEMPLATE
    # -----------------------------

    with open(template_file, "r", encoding="utf-8") as f:
        template_content = f.read()

    final_content = template_content
    final_content = final_content.replace("{{TABLE_ROWS}}", rows_html)
    final_content = final_content.replace("{{SUMMARY_PLACEHOLDER}}", summary_html)
    final_content = final_content.replace("{{EMPLOYEE_NAME}}", str(employee_name))
    final_content = final_content.replace("{{MONTH_YEAR}}", str(month_year))
    final_content = final_content.replace("chart1.png", chart1_url)
    final_content = final_content.replace("chart2.png", chart2_url)
    final_content = final_content.replace("ppmc_logo.jpg", logo1_url)
    final_content = final_content.replace("State_emblem_of_Pakistan.png", logo2_url)

    # -----------------------------
    # SAVE HTML
    # -----------------------------

    with open(final_html, "w", encoding="utf-8") as f:
        f.write(final_content)

    # -----------------------------
    # CONVERT TO PDF
    # -----------------------------

    os.system(
        f'wkhtmltopdf '
        f'--enable-local-file-access '
        f'--orientation Landscape '
        f'--page-size A4 '
        f'--margin-top 5mm '
        f'--margin-bottom 5mm '
        f'--margin-left 5mm '
        f'--margin-right 5mm '
        f'"{final_html}" "{final_pdf}"'
    )

    print(f"Report generated for {sheet_name}")


# =====================================================
# MAIN EXECUTION
# =====================================================

project_folder = r"C:\Users\ppmc\Desktop\ppmc"

excel_path = os.path.join(project_folder, "Attendance_Summary_Auto_Updated.xlsx")
template_path = os.path.join(project_folder, "report.md")
ppmc_logo = os.path.join(project_folder, "ppmc_logo.jpg")
state_logo = os.path.join(project_folder, "State_emblem_of_Pakistan.png")

desktop = os.path.join(os.path.expanduser("~"), "Desktop")
temp_folder = os.path.join(desktop, "temp")

if not os.path.exists(temp_folder):
    os.makedirs(temp_folder)

excel_data = pd.ExcelFile(excel_path)
all_sheets = excel_data.sheet_names

for sheet in all_sheets:
    generate_attendance_report(
        excel_path,
        template_path,
        ppmc_logo,
        state_logo,
        sheet,
        temp_folder
    )

print("All sheet reports generated successfully inside Desktop/temp folder.")
