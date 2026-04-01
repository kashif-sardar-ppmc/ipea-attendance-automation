# import pandas as pd
# import matplotlib.pyplot as plt
# import os

# # -----------------------------
# # PATH SETTINGS
# # -----------------------------

# project_folder = r"C:\Users\ppmc\Desktop\ppmc"

# input_file = os.path.join(project_folder, "Attendance_Summary_Auto_Updated.xlsx")
# template_file = os.path.join(project_folder, "report.md")

# final_html = os.path.join(project_folder, "final_report.html")
# final_pdf = os.path.join(project_folder, "Attendance_Report.pdf")

# chart1_path = os.path.join(project_folder, "chart1.png")
# chart2_path = os.path.join(project_folder, "chart2.png")

# sheet_name = "Sheet 21"
# # -----------------------------
# # LOAD DATA
# # -----------------------------

# df = pd.read_excel(input_file, sheet_name=sheet_name)
# df.columns = df.columns.str.strip()
# df = df.fillna("")

# # -----------------------------
# # EXTRACT EMPLOYEE NAME
# # -----------------------------

# employee_name = df["Name"].iloc[0]

# # -----------------------------
# # EXTRACT MONTH & YEAR
# # -----------------------------

# # Convert Date column to datetime safely
# df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

# # Drop invalid dates
# valid_dates = df["Date"].dropna()

# # Extract Month & Year safely
# if not valid_dates.empty:
#     month_year = valid_dates.iloc[0].strftime("%B %Y")
# else:
#     month_year = "N/A"

# # Format Date column for table display
# df["Date"] = df["Date"].dt.strftime("%d-%m-%Y")


# # -----------------------------
# # ROUND NUMERIC COLUMNS
# # -----------------------------

# numeric_cols = ["Worked Minutes", "Late Minutes", "Less Than 8 Hours"]

# for col in numeric_cols:
#     df[col] = pd.to_numeric(df[col], errors="coerce").round(0)

# # -----------------------------
# # CREATE COMPACT TABLE
# # -----------------------------

# display_columns = [
#     "Date",
#     "Time In",
#     "Time Out",
#     "Worked Minutes",
#     "Late Minutes",
#     "Less Than 8 Hours",
#     "Status",
#     "Comments"
# ]

# rows_html = ""

# for index, row in df.iterrows():
#     bg_color = "#f2f9f4" if index % 2 != 0 else "white"
#     rows_html += f"<tr style='background:{bg_color};'>"
    
#     for col in display_columns:
#         rows_html += f"<td style='padding:3px; border:1px solid #ccc; text-align:center;'>{row[col]}</td>"
    
#     rows_html += "</tr>"

# # -----------------------------
# # SUMMARY CALCULATIONS
# # -----------------------------

# total_days = df["Date"].nunique()
# present_days = df["Worked Minutes"].gt(0).sum()
# total_worked = df["Worked Minutes"].sum()
# total_late = df["Late Minutes"].sum()

# attendance_percentage = round((present_days / total_days) * 100, 2)

# summary_html = f"""
# <ul>
# <li><strong>Total Working Days:</strong> {total_days}</li>
# <li><strong>Days Present:</strong> {present_days}</li>
# <li><strong>Total Worked Minutes:</strong> {int(total_worked)}</li>
# <li><strong>Total Late Minutes:</strong> {int(total_late)}</li>
# <li><strong>Attendance Percentage:</strong> {attendance_percentage}%</li>
# </ul>
# """

# # -----------------------------
# # CREATE CHARTS
# # -----------------------------

# worked_numeric = pd.to_numeric(df["Worked Minutes"], errors="coerce")
# late_numeric = pd.to_numeric(df["Late Minutes"], errors="coerce")

# plt.figure()
# plt.plot(df["Date"], worked_numeric)
# plt.xticks(rotation=45)
# plt.tight_layout()
# plt.savefig(chart1_path)
# plt.close()

# plt.figure()
# plt.bar(df["Date"], late_numeric)
# plt.xticks(rotation=45)
# plt.tight_layout()
# plt.savefig(chart2_path)
# plt.close()

# # -----------------------------
# # LOAD TEMPLATE
# # -----------------------------

# with open(template_file, "r", encoding="utf-8") as f:
#     template_content = f.read()

# final_content = template_content.replace("{{TABLE_ROWS}}", rows_html)
# final_content = final_content.replace("{{SUMMARY_PLACEHOLDER}}", summary_html)
# final_content = final_content.replace("{{EMPLOYEE_NAME}}", str(employee_name))
# final_content = final_content.replace("{{MONTH_YEAR}}", str(month_year))

# # -----------------------------
# # SAVE MARKDOWN
# # -----------------------------

# with open(final_html, "w", encoding="utf-8") as f:
#     f.write(final_content)

# print("Markdown saved on Desktop.")

# # -----------------------------
# # CONVERT TO CLEAN LANDSCAPE PDF
# # -----------------------------

# os.system(
#     f'wkhtmltopdf '
#     f'--enable-local-file-access '
#     f'--orientation Landscape '
#     f'--page-size A4 '
#     f'--margin-top 5mm '
#     f'--margin-bottom 5mm '
#     f'--margin-left 5mm '
#     f'--margin-right 5mm '
#     f'--disable-smart-shrinking '
#     f'"{final_html}" "{final_pdf}"'
# )


# print("PDF saved on Desktop successfully.")



import pandas as pd
import matplotlib.pyplot as plt
import os

# -----------------------------
# PATH SETTINGS
# -----------------------------

project_folder = r"C:\Users\ppmc\Desktop\ppmc"

input_file = os.path.join(project_folder, "Attendance_Summary_Auto_Updated.xlsx")
template_file = os.path.join(project_folder, "report.md")

final_html = os.path.join(project_folder, "final_report.html")
final_pdf = os.path.join(project_folder, "Attendance_Report.pdf")

chart1_path = os.path.join(project_folder, "chart1.png")
chart2_path = os.path.join(project_folder, "chart2.png")

sheet_name = "Sheet 21"

# -----------------------------
# LOAD DATA
# -----------------------------

df = pd.read_excel(input_file, sheet_name=sheet_name)
df.columns = df.columns.str.strip()

# -----------------------------
# DATE HANDLING
# -----------------------------

df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

valid_dates = df["Date"].dropna()

if not valid_dates.empty:
    month_year = valid_dates.iloc[0].strftime("%B %Y")
else:
    month_year = "N/A"

# Format date for display
df["Date"] = df["Date"].dt.strftime("%d-%m-%Y")

# -----------------------------
# EXTRACT EMPLOYEE NAME
# -----------------------------

employee_name = df["Name"].iloc[0]

# -----------------------------
# ROUND NUMERIC COLUMNS
# -----------------------------

numeric_cols = ["Worked Minutes", "Late Minutes", "Less Than 8 Hours"]

for col in numeric_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce").round(0)

# Now clean empty values AFTER numeric processing
df = df.fillna("")

# -----------------------------
# CREATE TABLE WITH RED HIGHLIGHT
# -----------------------------

display_columns = [
    "Date",
    "Time In",
    "Time Out",
    "Worked Minutes",
    "Late Minutes",
    "Less Than 8 Hours",
    "Status",
    "Comments"
]

rows_html = ""

for index, row in df.iterrows():

    less_than_8 = pd.to_numeric(row["Less Than 8 Hours"], errors="coerce")

    # Highlight if worked less than 8 hours
    if pd.notna(less_than_8) and less_than_8 > 0:
        bg_color = "#ffdddd"
    else:
        bg_color = "#f2f9f4" if index % 2 != 0 else "white"

    rows_html += f"<tr style='background:{bg_color};'>"

    for col in display_columns:
        rows_html += f"<td style='padding:5px; border:1px solid #ccc; text-align:center; font-size:12px;'>{row[col]}</td>"

    rows_html += "</tr>"

# -----------------------------
# SUMMARY CALCULATIONS
# -----------------------------

# Convert again for accurate summary math
worked_numeric = pd.to_numeric(df["Worked Minutes"], errors="coerce")
late_numeric = pd.to_numeric(df["Late Minutes"], errors="coerce")

total_days = df["Date"].nunique()
present_days = worked_numeric.gt(0).sum()
total_worked = worked_numeric.sum()
total_late = late_numeric.sum()
late_days = df["Status"].astype(str).str.strip().str.lower().eq("late").sum()


attendance_percentage = round((present_days / total_days) * 100, 2) if total_days > 0 else 0

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

plt.figure(figsize=(6,4))
plt.plot(df["Date"], worked_numeric)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(chart1_path)
plt.close()

plt.figure(figsize=(6,4))
plt.bar(df["Date"], late_numeric)
plt.xticks(rotation=45)
plt.tight_layout()
plt.savefig(chart2_path)
plt.close()

# -----------------------------
# LOAD TEMPLATE & REPLACE
# -----------------------------

with open(template_file, "r", encoding="utf-8") as f:
    template_content = f.read()

final_content = template_content.replace("{{TABLE_ROWS}}", rows_html)
final_content = final_content.replace("{{SUMMARY_PLACEHOLDER}}", summary_html)
final_content = final_content.replace("{{EMPLOYEE_NAME}}", str(employee_name))
final_content = final_content.replace("{{MONTH_YEAR}}", str(month_year))

# -----------------------------
# SAVE FINAL HTML
# -----------------------------

with open(final_html, "w", encoding="utf-8") as f:
    f.write(final_content)

print("HTML generated successfully.")

# -----------------------------
# CONVERT HTML TO LANDSCAPE PDF
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
    f'--disable-smart-shrinking '
    f'"{final_html}" "{final_pdf}"'
)

print("PDF saved successfully.")

