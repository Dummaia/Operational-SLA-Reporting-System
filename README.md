# 🚀 Operational SLA Reporting System

Python automation for monitoring operational activities and automatically delivering executive and individual SLA reports.

This project was inspired by a real-world operational automation developed to improve daily decision-making. It consolidates operational data from multiple Excel worksheets, applies business rules, identifies critical situations and automatically delivers customized reports to managers and operational users.

The goal is to transform operational data into actionable information, ensuring that the right people receive the right information at the right time.

---

# ✨ Key Features

- 📥 Read operational data from multiple Excel worksheets
- 🔄 Consolidate operational information
- 🧹 Clean and validate datasets
- 📊 Apply business rules
- 🚨 Detect overdue operational activities
- ⚠️ Identify upcoming SLA risks
- 📝 Detect outdated operational updates
- 📈 Generate Executive SLA Reports
- 👤 Generate Individual SLA Reports
- 📧 Automatically deliver Outlook email reports

---

# 🏗 Architecture

```
                  Operational Data
                         │
                         ▼
              Business Rules Engine
                         │
        ┌────────────────┴────────────────┐
        │                                 │
        ▼                                 ▼
 Executive SLA Report          Individual SLA Reports
        │                                 │
        ▼                                 ▼
    Management                 Responsible Users
```

---

# 📂 Repository Structure

```text
Operational-Intelligence-Suite/

│
├── README.md
├── LICENSE
├── requirements.txt
├── .gitignore
│
├── main_executive_report.py
├── main_individual_report.py
│
├── data/
│   ├── sample_operations.xlsx
│   └── sample_users.xlsx
│
└── images/
```

---

# ⚙ Components

## 📊 Executive SLA Report

Generates a consolidated management report containing:

- Executive operational KPIs
- Open operational activities
- Overdue activities
- Immediate SLA risks
- Upcoming SLA attention
- Critical workflow stages
- Outdated operational updates
- Top critical responsible users
- Automatic Outlook delivery

---

## 👤 Individual SLA Report

Generates personalized reports for each responsible user containing:

- Assigned operational activities
- Overdue items
- Upcoming deadlines
- Current workflow stage
- Latest operational updates
- Prioritized action list
- Automatic Outlook delivery

---

# 🛠 Technologies

- 🐍 Python
- 📊 Pandas
- 🔢 NumPy
- 📄 OpenPyXL
- 📧 PyWin32 (Outlook Automation)
- 📂 Excel

---

# ⚙ Workflow

```
Excel Worksheets
        │
        ▼
Read Multiple Worksheets
        │
        ▼
Data Cleaning & Validation
        │
        ▼
Business Rules Engine
        │
        ├───────────────┐
        │               │
        ▼               ▼
Executive Report   Individual Reports
        │               │
        ▼               ▼
Management     Responsible Users
```

---

# 🎯 Purpose

This project was not created to replace dashboards.

Dashboards are essential for monitoring performance and analyzing operational indicators.

The purpose of this solution is different.

It transforms existing operational information into proactive communication, helping managers and operational teams identify priorities, mitigate SLA risks and keep operational information up to date before issues become critical.

---

# 🚀 Future Improvements

- PostgreSQL integration
- Historical SLA database
- Power BI integration
- Microsoft Teams notifications
- Docker deployment
- Azure Container Apps
- REST API
- Web Dashboard
- AI-assisted operational insights

---

# 📸 Screenshots

The repository includes examples of:

- Executive SLA email
- Individual SLA email
- Operational reports

---

# 📄 License

This project is licensed under the MIT License.

---

# 👨‍💻 Author

**Eduardo Maia**

Data Analytics | Business Intelligence | Process Automation

GitHub:
https://github.com/maiae3381
