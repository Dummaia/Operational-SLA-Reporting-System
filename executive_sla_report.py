"""
=========================================================
Operational SLA Reporting System - Executive Report
---------------------------------------------------------

Python automation for monitoring operational activities,
applying business rules and automatically delivering
executive SLA reports.

Author:
Eduardo Maia

Version:
1.0.0

License:
MIT License

=========================================================
"""

import pandas as pd
import numpy as np
import win32com.client as win32
from pathlib import Path
from datetime import datetime

__author__ = "Eduardo Maia"
__version__ = "1.0.0"
__license__ = "MIT"

# =========================
# CONFIGURATIONS
# =========================

# 2. Alterado para "open" por padrão para evitar envios acidentais
SEND_MODE = "open"  # "open" for testing/review / "send" for production

PROJECT_FOLDER = Path(__file__).parent
DATA_FOLDER = PROJECT_FOLDER / "data"

OPERATIONS_FILE = DATA_FOLDER / "sample_operations.xlsx"
USERS_FILE = DATA_FOLDER / "sample_users.xlsx"

# 1. & 5. Nomenclatura atualizada de EXPIRED para OVERDUE
WORKSHEETS = [
    "OVERDUE",
    "0 TO 7 DAYS",
    "8 TO 15 DAYS",
    "16 TO 25 DAYS",
    "26 TO 35 DAYS"
]

# =========================
# READ DATASET
# =========================

base_dataframes = []

for worksheet in WORKSHEETS:
    temp_df = pd.read_excel(OPERATIONS_FILE, sheet_name=worksheet)
    temp_df.columns = temp_df.columns.str.strip()
    temp_df["ORIGIN_WORKSHEET"] = worksheet
    base_dataframes.append(temp_df)

operations_df = pd.concat(base_dataframes, ignore_index=True)
operations_df.columns = operations_df.columns.str.strip()

# =========================
# DATA CLEANING
# =========================

operations_df["OPERATION_TYPE"] = operations_df["OPERATION_TYPE"].astype(str).str.strip().str.upper()
operations_df["SLA_STATUS"] = operations_df["SLA_STATUS"].astype(str).str.strip().str.upper()
operations_df["SLA_RANGE"] = operations_df["SLA_RANGE"].astype(str).str.strip()
operations_df["RESPONSIBLE_USER"] = operations_df["RESPONSIBLE_USER"].astype(str).str.strip().str.upper()

operations_df["CURRENT_STAGE"] = operations_df["CURRENT_STAGE"].replace("", pd.NA)
operations_df["LATEST_UPDATE"] = operations_df["LATEST_UPDATE"].replace("", pd.NA)

operations_df["SLA_DAYS_REMAINING"] = pd.to_numeric(
    operations_df["SLA_DAYS_REMAINING"],
    errors="coerce"
)

operations_df["REQUEST_ID"] = pd.to_numeric(
    operations_df["REQUEST_ID"],
    errors="coerce"
)

operations_df = operations_df[operations_df["LINE_ID"].notna()]
operations_df = operations_df[operations_df["OPERATION_TYPE"].isin(["STANDARD", "BACKLOG"])]
operations_df = operations_df.drop_duplicates(subset=["LINE_ID"])

operations_df["REQUEST_ID"] = operations_df["REQUEST_ID"].astype("Int64")
operations_df["SLA_DAYS_REMAINING"] = operations_df["SLA_DAYS_REMAINING"].astype("Int64")

# =========================
# OUTDATED OPERATIONAL UPDATES - BUSINESS DAYS
# =========================

def extract_update_date(text_content):
    if pd.isna(text_content):
        return pd.NaT

    text_content = str(text_content).strip()
    date_text = text_content[:5]

    # 8. except Exception: no lugar de except:
    try:
        date_with_year = date_text + f"/{datetime.now().year}"
        return pd.to_datetime(
            date_with_year,
            format="%d/%m/%Y",
            errors="coerce"
        )
    except Exception:
        return pd.NaT


def calculate_business_days(update_date):
    if pd.isna(update_date):
        return pd.NA

    return np.busday_count(
        update_date.date(),
        datetime.now().date()
    )


operations_df["LAST_UPDATE_DATE"] = operations_df["LATEST_UPDATE"].apply(extract_update_date)
operations_df["BUSINESS_DAYS_WITHOUT_UPDATE"] = operations_df["LAST_UPDATE_DATE"].apply(
    calculate_business_days
)

outdated_updates_df = operations_df[
    (operations_df["BUSINESS_DAYS_WITHOUT_UPDATE"] > 3) &
    (operations_df["LATEST_UPDATE"].notna())
].copy()

# =========================
# INDICATORS / KPIS
# =========================

total_open_items = len(operations_df)

standard_type = operations_df[operations_df["OPERATION_TYPE"] == "STANDARD"]
backlog_type = operations_df[operations_df["OPERATION_TYPE"] == "BACKLOG"]

overdue_items = operations_df[operations_df["SLA_STATUS"] == "OVERDUE"]
on_time_items = operations_df[operations_df["SLA_STATUS"] == "ON TIME"]

overdue_standard = overdue_items[overdue_items["OPERATION_TYPE"] == "STANDARD"]
overdue_backlog = overdue_items[overdue_items["OPERATION_TYPE"] == "BACKLOG"]

risk_0_5_days = operations_df[
    (operations_df["SLA_DAYS_REMAINING"] >= 0) &
    (operations_df["SLA_DAYS_REMAINING"] <= 5)
].copy()

warning_6_7_days = operations_df[
    (operations_df["SLA_DAYS_REMAINING"] >= 6) &
    (operations_df["SLA_DAYS_REMAINING"] <= 7)
]

missing_update_items = operations_df[operations_df["CURRENT_STAGE"].isna()]

critical_users = (
    overdue_items.groupby("RESPONSIBLE_USER")
    .size()
    .sort_values(ascending=False)
    .head(5)
)

overdue_stages = (
    overdue_items["CURRENT_STAGE"]
    .fillna("NO UPDATE")
    .value_counts()
    .head(10)
)

outdated_updates_by_user = (
    outdated_updates_df.groupby("RESPONSIBLE_USER")
    .size()
    .sort_values(ascending=False)
)

# =========================
# TABLES
# =========================

# 10. Garante exibição da coluna ITEM
risk_cols = ["REQUEST_ID", "ITEM", "RESPONSIBLE_USER", "SLA_DAYS_REMAINING", "ITEM_TYPE", "CURRENT_STAGE"]
available_risk_cols = [c for c in risk_cols if c in risk_0_5_days.columns]

risk_table_df = (
    risk_0_5_days[available_risk_cols]
    .sort_values("SLA_DAYS_REMAINING")
    .head(15)
    .rename(columns={
        "REQUEST_ID": "ID",
        "ITEM": "Item",
        "RESPONSIBLE_USER": "Responsible User",
        "SLA_DAYS_REMAINING": "Days Remaining",
        "ITEM_TYPE": "Item Type",
        "CURRENT_STAGE": "Current Stage"
    })
)

outdated_cols = ["REQUEST_ID", "ITEM", "RESPONSIBLE_USER", "BUSINESS_DAYS_WITHOUT_UPDATE", "CURRENT_STAGE", "LATEST_UPDATE"]
available_outdated_cols = [c for c in outdated_cols if c in outdated_updates_df.columns]

outdated_updates_table_df = (
    outdated_updates_df[available_outdated_cols]
    .sort_values("BUSINESS_DAYS_WITHOUT_UPDATE", ascending=False)
    .head(15)
    .rename(columns={
        "REQUEST_ID": "ID",
        "ITEM": "Item",
        "RESPONSIBLE_USER": "Responsible User",
        "BUSINESS_DAYS_WITHOUT_UPDATE": "Business Days",
        "CURRENT_STAGE": "Current Stage",
        "LATEST_UPDATE": "Latest Update"
    })
)

def render_html_table(table_df):
    if table_df.empty:
        return "<p>No records found.</p>"

    # 7. fillna("") para tratamento seguro de células vazias
    html_code = table_df.fillna("").to_html(
        index=False,
        border=0,
        justify="left"
    )

    html_code = html_code.replace(
        '<table class="dataframe">',
        '<table style="border-collapse:collapse; width:100%; font-family:Segoe UI, Arial, sans-serif; font-size:13px; margin-top:8px; margin-bottom:18px;">'
    )

    html_code = html_code.replace(
        "<thead>",
        '<thead style="background-color:#f3f4f6;">'
    )

    html_code = html_code.replace(
        "<th>",
        '<th style="border:1px solid #d1d5db; padding:7px; text-align:left; font-weight:600;">'
    )

    html_code = html_code.replace(
        "<td>",
        '<td style="border:1px solid #d1d5db; padding:7px; text-align:left;">'
    )

    return html_code

# =========================
# TEXT FORMATTING
# =========================

execution_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")

critical_users_html = ""

for index, (responsible_user, quantity) in enumerate(critical_users.items(), start=1):
    critical_users_html += f"<p>{index}º {responsible_user}: <b>{quantity}</b> overdue</p>"

overdue_stages_html = ""

for stage_name, quantity in overdue_stages.items():
    overdue_stages_html += f"<p>{stage_name}: <b>{quantity}</b></p>"

user_updates_html = ""

if len(outdated_updates_by_user) > 0:
    for responsible_user, quantity in outdated_updates_by_user.items():
        user_updates_html += f"<p>{responsible_user}: <b>{quantity}</b></p>"
else:
    user_updates_html = "<p>No outdated operational updates found.</p>"

# =========================
# EMAIL HTML TEMPLATE
# =========================

# 3. Alterado de "Please find attached" para "Please find below"
# 4. Assinatura atualizada para "Operational SLA Reporting System – Eduardo Maia"
management_email_html = f"""
<html>
<head>
<style>
    body {{
        font-family: Segoe UI, Arial, sans-serif;
        font-size: 14px;
        color: #222222;
        line-height: 1.5;
    }}

    h2 {{
        margin-bottom: 4px;
        color: #1f2937;
    }}

    h3 {{
        margin-top: 26px;
        border-bottom: 1px solid #d1d5db;
        padding-bottom: 6px;
        color: #1f2937;
    }}

    .summary {{
        border-collapse: collapse;
        width: 760px;
        margin-top: 10px;
        margin-bottom: 18px;
        table-layout: fixed;
    }}

    .summary td {{
        padding: 9px 8px;
        border-bottom: 1px solid #d1d5db;
        vertical-align: middle;
    }}

    .summary td:nth-child(1) {{
        width: 380px;
        font-weight: 600;
    }}

    .summary td:nth-child(2) {{
        width: 80px;
        text-align: center;
        font-weight: 700;
    }}

    .summary td:nth-child(3) {{
        width: 300px;
        white-space: nowrap;
    }}

    .signature {{
        margin-top: 28px;
    }}

    .brand {{
        color: #374151;
        font-weight: 600;
    }}
</style>
</head>

<body>

<p>Hello,</p>

<p>Please find below the updated management report for the operational portfolio.</p>

<h2>Management Report - SLA Alert</h2>
<p><b>Execution Date:</b> {execution_timestamp}</p>

<h3>Executive Summary</h3>

<table class="summary">
    <tr>
        <td>📦 Total Open Lines</td>
        <td>{total_open_items}</td>
        <td>STANDARD: <b>{len(standard_type)}</b> | BACKLOG: <b>{len(backlog_type)}</b></td>
    </tr>

    <tr>
        <td>🔴 Overdue</td>
        <td>{len(overdue_items)}</td>
        <td>STANDARD: <b>{len(overdue_standard)}</b> | BACKLOG: <b>{len(overdue_backlog)}</b></td>
    </tr>

    <tr>
        <td>🟡 Attention 6 to 7 Days</td>
        <td>{len(warning_6_7_days)}</td>
        <td></td>
    </tr>

    <tr>
        <td>🟠 Immediate Risk 0 to 5 Days</td>
        <td>{len(risk_0_5_days)}</td>
        <td></td>
    </tr>

    <tr>
        <td>🟢 On Time</td>
        <td>{len(on_time_items)}</td>
        <td></td>
    </tr>

    <tr>
        <td>📋 Requisitions without Operational Update</td>
        <td>{len(missing_update_items)}</td>
        <td></td>
    </tr>

    <tr>
        <td>⏳ Outdated Operational Update &gt; 3 Business Days</td>
        <td>{len(outdated_updates_df)}</td>
        <td></td>
    </tr>
</table>

<h3>Most Critical Responsible Users</h3>
{critical_users_html}

<h3>Main Overdue Process Stages</h3>
{overdue_stages_html}

<h3>Immediate Risk - 0 to 5 Days</h3>
{render_html_table(risk_table_df)}

<h3>Outdated Operational Update - More than 3 Business Days</h3>

<p><b>Total Outdated Operational Updates:</b> {len(outdated_updates_df)}</p>

<p><b>By Responsible User:</b></p>
{user_updates_html}

<p><b>Detailed Breakdown:</b></p>
{render_html_table(outdated_updates_table_df)}

<div class="signature">
    <p>Best regards,</p>
    <p class="brand">Operational SLA Reporting System – Eduardo Maia</p>
</div>

</body>
</html>
"""

# =========================
# FETCH USER EMAILS
# =========================

users_df = pd.read_excel(USERS_FILE)
users_df.columns = users_df.columns.str.strip()
users_df["RESPONSIBLE_USER"] = users_df["RESPONSIBLE_USER"].astype(str).str.strip().str.upper()

# 7. Preenchimento de vazios antes do filtro
users_df["EMAIL"] = users_df["EMAIL"].fillna("").astype(str).str.strip()

# 10. Filtro Dinâmico: Busca por coluna ROLE ou MANAGEMENT para evitar nomes hardcoded no código
if "ROLE" in users_df.columns:
    management_users = users_df[users_df["ROLE"].astype(str).str.upper() == "MANAGEMENT"]
elif "IS_MANAGEMENT" in users_df.columns:
    management_users = users_df[users_df["IS_MANAGEMENT"] == True]
else:
    # Fallback caso a coluna não exista: busca por cargos de gestão no campo ROLE/POSITION
    management_users = users_df[
        users_df["RESPONSIBLE_USER"].str.contains("MANAGEMENT|DIRECTOR|MANAGER", na=False)
    ]

# Se o filtro dinâmico não encontrar nada, recupera e-mails válidos padrão
if management_users.empty:
    management_users = users_df[
        (users_df["EMAIL"] != "") & (users_df["EMAIL"].str.lower() != "nan")
    ]

destination_emails = management_users["EMAIL"].dropna().tolist()
recipient_list = "; ".join(destination_emails)

print("Recipients found:")
print(recipient_list)

# =========================
# OUTLOOK INTEGRATION
# =========================

# 8. Tratamento com except Exception:
try:
    outlook_app = win32.Dispatch("Outlook.Application")
    email_message = outlook_app.CreateItem(0)

    email_message.To = recipient_list
    email_message.Subject = f"Management Report - SLA Alert - {datetime.now().strftime('%d/%m/%Y')}"
    email_message.HTMLBody = management_email_html

    if SEND_MODE == "send":
        email_message.Send()
        print("Email sent successfully.")
    else:
        email_message.Display()
        print("Email opened for verification.")
except Exception as e:
    print(f"Failed to deliver executive report: {e}")

input("\nPress Enter to finish...")