<div style="font-family:Segoe UI, Arial; margin:10px;">

<table style="width:100%; border-bottom:4px solid #1f6f3e; padding-bottom:8px;">
<tr>
<td style="width:15%; text-align:left;"><img src="ppmc_logo.jpg" width="90"></td>
<td style="width:70%; text-align:center;">
    <h2 style="color:#1f6f3e; margin:0; font-size:22px;">Executive Attendance Report</h2>
    <p style="margin:4px 0; color:#4caf50; font-size:16px;">Power Planning & Monitoring Company</p>
    <p style="margin:0; font-weight:bold;">Reporting Period: {{MONTH_YEAR}}</p>
</td>
<td style="width:15%; text-align:right;"><img src="State_emblem_of_Pakistan.png" width="80"></td>
</tr>
</table>

<h3 style="background:#1f6f3e; color:white; padding:8px; border-radius:3px; font-size:15px; margin-top:20px;">
Workforce Productivity Overview
</h3>
<div style="border:1px solid #ccc; padding:15px; background:#f9f9f9; border-radius:5px;">
{{EXECUTIVE_SUMMARY}}
</div>

<table style="width:100%; margin-top:20px;">
<tr>
    <td style="width:50%; text-align:center;">
        <strong style="font-size:12px;">Workforce Presence vs Absence</strong><br>
        <img src="exec_chart1.png" style="width:95%;">
    </td>
    <td style="width:50%; text-align:center;">
        <strong style="font-size:12px;">Top 5 Performers (Avg. Hours)</strong><br>
        <img src="exec_chart2.png" style="width:95%;">
    </td>
</tr>
</table>

<h3 style="background:#1f6f3e; color:white; padding:8px; border-radius:3px; font-size:15px; margin-top:20px;">
Individual Employee Performance Summary
</h3>
<table style="width:100%; border-collapse:collapse; font-size:11px;">
<tr style="background:#1f6f3e; color:white; text-align:center;">
    <th style="padding:8px; text-align:left;">Employee Name</th>
    <th style="padding:8px;">Department</th>
    <th style="padding:8px;">Days Present</th>
    <th style="padding:8px;">Days Absent</th>
    <th style="padding:8px;">Under 8h Days</th>
    <th style="padding:8px;">Avg. Hours/Day</th>
    <th style="padding:8px;">Status</th>
</tr>
{{EXECUTIVE_TABLE_ROWS}}
</table>

<div style="text-align:center; font-size:10px; color:#777; margin-top:30px;">
Generated for Management Review | PPMC Confidential
</div>
</div>