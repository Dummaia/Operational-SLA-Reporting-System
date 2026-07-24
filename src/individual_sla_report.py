"""
=========================================================
Operational SLA Reporting System - Individual Report
---------------------------------------------------------

Python automation for monitoring operational activities,
applying business rules and automatically delivering
individual SLA reports.

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

# 2. Alterado para "open" por padrão para maior segurança
SEND_MODE = "open"  # "open" for testing/review / "send" for production

# To test only one specific user:
# TEST_RESPONSIBLE_USER = "JOHN SMITH"

# 6. Comentário ajustado para ficar mais natural
# Process all responsible users
TEST_RESPONSIBLE_USER = None

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
# READ OPERATIONS DATASET
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
# OPERATIONS DATA CLEANING
# =========================

operations_df["OPERATION_TYPE"] = operations_df["OPERATION_TYPE"].astype(str).str.strip().str.upper()
operations_df["SLA_STATUS"] = operations_df["SLA_STATUS"].astype(str).str.strip().str.upper()
operations_df["SLA_RANGE"] = operations_df["SLA_RANGE"].astype(str).str.strip()
operations_df["RESPONSIBLE_USER"] = operations_df["RESPONSIBLE_USER"].astype(str).str.strip().str.upper()

operations_df["CURRENT_STAGE"] = operations_df["CURRENT_STAGE"].replace("", pd.NA)
operations_df["LATEST_UPDATE"] = operations_df["LATEST_UPDATE"].replace("", pd.NA)

operations_df["SLA_DAYS_REMAINING"] = pd.to_numeric(operations_df["SLA_DAYS_REMAINING"], errors="coerce")
operations_df["REQUEST_ID"] = pd.to_numeric(operations_df["REQUEST_ID"], errors="coerce")

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

    # 8. exception handling genérico trocado por except Exception:
    try:
        date_with_year = date_text + f"/{datetime.now().year}"
        return pd.to_datetime(date_with_year, format="%d/%m/%Y", errors="coerce")
    except Exception:
        return pd.NaT


def calculate_business_days(update_date):
    if pd.isna(update_date):
        return pd.NA

    return np.busday_count(update_date.date(), datetime.now().date())


operations_df["LAST_UPDATE_DATE"] = operations_df["LATEST_UPDATE"].apply(extract_update_date)
operations_df["BUSINESS_DAYS_WITHOUT_UPDATE"] = operations_df["LAST_UPDATE_DATE"].apply(calculate_business_days)

# =========================
# HTML FUNCTIONS
# =========================

def render_html_table(table_df):
    if table_df.empty:
        return "<p>No records found.</p>"

    # 7. fillna("") adicionado para garantir robustness com células nulas
    html_code = table_df.fillna("").to_html(index=False, border=0, justify="left")

    html_code = html_code.replace(
        '<table class="dataframe">',
        '<table style="border-collapse:collapse; width:100%; font-family:Segoe UI, Arial, sans-serif; font-size:13px; margin-top:8px; margin-bottom:18px;">'
    )

    html_code = html_code.replace(
        "<th>",
        '<th style="border:1px solid #d1d5db; background-color:#f3f4f6; padding:7px; text-align:left;">'
    )

    html_code = html_code.replace(
        "<td>",
        '<td style="border:1px solid #d1d5db; padding:7px; text-align:left;">'
    )

    return html_code


def generate_user_html_report(user_operations_df, responsible_user):
    total_open_items = len(user_operations_df)

    standard_type = user_operations_df[user_operations_df["OPERATION_TYPE"] == "STANDARD"]
    backlog_type = user_operations_df[user_operations_df["OPERATION_TYPE"] == "BACKLOG"]

    overdue_items = user_operations_df[user_operations_df["SLA_STATUS"] == "OVERDUE"]
    on_time_items = user_operations_df[user_operations_df["SLA_STATUS"] == "ON TIME"]

    risk_0_5_days = user_operations_df[
        (user_operations_df["SLA_DAYS_REMAINING"] >= 0) &
        (user_operations_df["SLA_DAYS_REMAINING"] <= 5)
    ].copy()

    warning_6_7_days = user_operations_df[
        (user_operations_df["SLA_DAYS_REMAINING"] >= 6) &
        (user_operations_df["SLA_DAYS_REMAINING"] <= 7)
    ].copy()

    missing_update_items = user_operations_df[user_operations_df["CURRENT_STAGE"].isna()].copy()

    outdated_updates_df = user_operations_df[
        (user_operations_df["BUSINESS_DAYS_WITHOUT_UPDATE"] > 3) &
        (user_operations_df["LATEST_UPDATE"].notna())
    ].copy()

    overdue_stages = (
        overdue_items["CURRENT_STAGE"]
        .fillna("NO UPDATE")
        .value_counts()
        .head(10)
    )

    # 10. Garante exibição da coluna ITEM no relatório
    standard_columns = [
        "REQUEST_ID",
        "ITEM",
        "BUSINESS_UNIT",
        "SLA_DAYS_REMAINING",
        "ITEM_TYPE",
        "CURRENT_STAGE"
    ]
    available_cols = [c for c in standard_columns if c in user_operations_df.columns]

    overdue_table_df = (
        overdue_items[available_cols]
        .sort_values("SLA_DAYS_REMAINING")
        .head(15)
        .rename(columns={
            "REQUEST_ID": "ID",
            "ITEM": "Item",
            "BUSINESS_UNIT": "Business Unit",
            "SLA_DAYS_REMAINING": "Days",
            "ITEM_TYPE": "Item Type",
            "CURRENT_STAGE": "Current Stage"
        })
    )

    risk_table_df = (
        risk_0_5_days[available_cols]
        .sort_values("SLA_DAYS_REMAINING")
        .head(15)
        .rename(columns={
            "REQUEST_ID": "ID",
            "ITEM": "Item",
            "BUSINESS_UNIT": "Business Unit",
            "SLA_DAYS_REMAINING": "Days",
            "ITEM_TYPE": "Item Type",
            "CURRENT_STAGE": "Current Stage"
        })
    )

    warning_table_df = (
        warning_6_7_days[available_cols]
        .sort_values("SLA_DAYS_REMAINING")
        .head(15)
        .rename(columns={
            "REQUEST_ID": "ID",
            "ITEM": "Item",
            "BUSINESS_UNIT": "Business Unit",
            "SLA_DAYS_REMAINING": "Days",
            "ITEM_TYPE": "Item Type",
            "CURRENT_STAGE": "Current Stage"
        })
    )

    outdated_cols = ["REQUEST_ID", "ITEM", "BUSINESS_DAYS_WITHOUT_UPDATE", "CURRENT_STAGE", "LATEST_UPDATE"]
    available_outdated_cols = [c for c in outdated_cols if c in user_operations_df.columns]

    outdated_updates_table_df = (
        outdated_updates_df[available_outdated_cols]
        .sort_values("BUSINESS_DAYS_WITHOUT_UPDATE", ascending=False)
        .head(15)
        .rename(columns={
            "REQUEST_ID": "ID",
            "ITEM": "Item",
            "BUSINESS_DAYS_WITHOUT_UPDATE": "Business Days",
            "CURRENT_STAGE": "Current Stage",
            "LATEST_UPDATE": "Latest Update"
        })
    )

    stages_text = ""

    if len(overdue_stages) > 0:
        for stage_name, quantity in overdue_stages.items():
            stages_text += f"<p>{stage_name}: <b>{quantity}</b></p>"
    else:
        stages_text = "<p>No overdue items found.</p>"

    execution_timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")

    # 3. Alterado de "Please find attached" para "Please find below"
    # 4. Assinatura atualizada para "Operational SLA Reporting System – Eduardo Maia"
    html_content = f"""
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
        color: #1f2937;
        margin-bottom: 4px;
    }}

    h3 {{
        margin-top: 26px;
        border-bottom: 1px solid #d1d5db;
        padding-bottom: 6px;
        color: #1f2937;
    }}

    .summary {{
        border-collapse: collapse;
        width: 720px;
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
        width: 390px;
        font-weight: 600;
    }}

    .summary td:nth-child(2) {{
        width: 80px;
        text-align: center;
        font-weight: 700;
    }}

    .summary td:nth-child(3) {{
        width: 250px;
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

<p>Hello, {responsible_user.title()}.</p>

<p>Please find below the SLA alert for your assigned portfolio.</p>

<h2>Individual SLA Alert</h2>
<p><b>Execution Date:</b> {execution_timestamp}</p>

<h3>Portfolio Summary</h3>

<table class="summary">
    <tr>
        <td>📦 Total Open Lines</td>
        <td>{total_open_items}</td>
        <td>STANDARD: <b>{len(standard_type)}</b> | BACKLOG: <b>{len(backlog_type)}</b></td>
    </tr>

    <tr>
        <td>🔴 Overdue</td>
        <td>{len(overdue_items)}</td>
        <td></td>
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

<h3>My Overdue Items</h3>
{render_html_table(overdue_table_df)}

<h3>Immediate Risk - 0 to 5 Days</h3>
{render_html_table(risk_table_df)}

<h3>Attention - 6 to 7 Days</h3>
{render_html_table(warning_table_df)}

<h3>Outdated Operational Update - More than 3 Business Days</h3>
{render_html_table(outdated_updates_table_df)}

<h3>Main Stages of My Overdue Items</h3>
{stages_text}

<div class="signature">
    <p>Best regards,</p>
    <p class="brand">Operational SLA Reporting System – Eduardo Maia</p>
</div>

</body>
</html>
"""

    return html_content


# =========================
# READ USERS
# =========================

users_df = pd.read_excel(USERS_FILE)
users_df.columns = users_df.columns.str.strip()

users_df["RESPONSIBLE_USER"] = users_df["RESPONSIBLE_USER"].astype(str).str.strip().str.upper()

# 7. Tratamento robusto para nulos com .fillna("")
users_df["EMAIL"] = users_df["EMAIL"].fillna("").astype(str).str.strip()

# Filter out rows without valid email addresses
users_df = users_df[
    (users_df["EMAIL"] != "") &
    (users_df["EMAIL"].str.lower() != "nan")
]

print()
print("Users loaded from file:")
print(users_df[["RESPONSIBLE_USER", "EMAIL"]])
print()

# Test mode: filter only for TEST_RESPONSIBLE_USER if provided
if TEST_RESPONSIBLE_USER is not None:
    users_df = users_df[users_df["RESPONSIBLE_USER"] == TEST_RESPONSIBLE_USER]

print("Users after test filter:")
print(users_df[["RESPONSIBLE_USER", "EMAIL"]])
print()

if users_df.empty:
    print("No target users found for delivery.")
    input("\nPress Enter to finish...")
    raise SystemExit

# =========================
# SEND EMAILS
# =========================

outlook_app = win32.Dispatch("Outlook.Application")

total_processed = 0
total_missing_portfolio = 0

print("Starting email generation process...")
print()

for _, row in users_df.iterrows():
    responsible_user = row["RESPONSIBLE_USER"]
    recipient_email = row["EMAIL"]

    user_operations_df = operations_df[operations_df["RESPONSIBLE_USER"] == responsible_user].copy()

    if user_operations_df.empty:
        print(f"No active portfolio found for: {responsible_user}")
        total_missing_portfolio += 1
        continue

    html_content = generate_user_html_report(user_operations_df, responsible_user)

    # 8. exception handling com except Exception:
    try:
        email_message = outlook_app.CreateItem(0)
        email_message.To = recipient_email
        email_message.Subject = f"Individual SLA Alert - {responsible_user.title()} - {datetime.now().strftime('%d/%m/%Y')}"
        email_message.HTMLBody = html_content

        if SEND_MODE == "send":
            email_message.Send()
            print(f"Email successfully sent to {responsible_user} -> {recipient_email}")
        else:
            email_message.Display()
            print(f"Email opened for verification: {responsible_user} -> {recipient_email}")

        total_processed += 1
    except Exception as e:
        print(f"Failed to process email for {responsible_user}: {e}")

print()
print("Process completed.")
print(f"Total emails processed: {total_processed}")
print(f"Users without portfolio data: {total_missing_portfolio}")

input("\nPress Enter to finish...")