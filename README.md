# 📋 Inspection Data Collection System

A secure, Flask-based web application for digitizing equipment inspection workflows. The platform replaces manual paper-based processes with a centralized web interface, allowing multiple users to collect, validate, and manage inspection records while maintaining synchronized data storage and real-time progress tracking.

> **Portfolio Project:** This repository has been generalized to showcase the software architecture and engineering concepts. Organization-specific data, infrastructure details, and sensitive configuration have been removed.

---

## 🚀 Overview

The Inspection Data Collection System streamlines operational data collection by replacing manual inspection forms with a centralized digital workflow.

Users can quickly locate an asset, complete standardized inspection forms, submit results, and monitor overall progress through a live dashboard. The application supports multiple concurrent users while ensuring data integrity through synchronized write operations.

---

## ✨ Features

- 🔍 Asset lookup using a unique equipment identifier
- 📋 Digital inspection forms with built-in validation
- ⚡ Automatic retrieval of asset information from a master inventory
- 📊 Real-time dashboard with progress tracking
- 👥 Multi-user support across a local network
- 🔒 Session-based authentication
- 📁 Centralized Excel-based data management
- ✅ Thread-safe concurrent data saving
- 🎨 Color-coded inspection status indicators
- 🔐 Optional HTTPS support for secure local deployments

---

## 🏗️ System Architecture

```text
                +----------------------+
                |   Web Browser        |
                +----------+-----------+
                           |
                           |
                    Flask Web Server
                           |
        +------------------+------------------+
        |                  |                  |
        |                  |                  |
 Authentication      Inspection Forms    Dashboard
        |                  |                  |
        +------------------+------------------+
                           |
                   Data Management Layer
                           |
                 Shared Excel Workbook
```

---

## 🖥️ Application Workflow

1. User logs into the application.
2. Enter or scan an equipment identifier.
3. Asset information is automatically retrieved.
4. Complete the digital inspection checklist.
5. Submit inspection results.
6. Data is safely written to the shared workbook.
7. Dashboard updates automatically with the latest progress.

---

## 🔒 Security Improvements

Several improvements were implemented to strengthen the original application design:

- Session-based authentication across protected routes
- Configurable application credentials
- Removal of hard-coded user information
- Secure handling of local configuration files
- Improved application initialization
- Separation of sensitive deployment assets from source control

Additional implementation details are available in **SECURITY.md**.

---

## 💻 Technology Stack

### Backend

- Python
- Flask

### Data Processing

- OpenPyXL
- Excel-based storage

### Security

- Session Authentication
- HTTPS Support
- Environment Variables

### Frontend

- HTML
- CSS
- JavaScript

---

## 📂 Project Structure

```text
inspection-data-collection/
│
├── README.md
├── SECURITY.md
├── LICENSE
├── requirements.txt
├── .gitignore
│
├── app.py
├── data_store.py
├── staff.example.json
│
├── templates/
│   ├── login.html
│   ├── form.html
│   ├── update.html
│   └── dashboard.html
│
├── static/
│
└── scripts/
    ├── START.bat
    ├── STOP_APP.bat
    ├── RESTART_APP.bat
    └── GENERATE_CERT.bat
```

---

## ⚙️ Installation

Clone the repository

```bash
git clone https://github.com/yourusername/inspection-data-collection.git
```

Install dependencies

```bash
pip install -r requirements.txt
```

Create your local configuration

```text
staff.example.json → staff.json
```

Start the application

```bash
python app.py
```

---

## 👥 Multi-User Support

The system is designed for collaborative environments where multiple users can simultaneously submit inspection records.

Thread-safe write operations ensure that concurrent updates do not corrupt the shared data source.

---

## 📈 Dashboard

The built-in dashboard provides:

- Overall completion progress
- Inspection statistics
- Status summaries
- Live updates
- Operational reporting

---

## 🎯 Skills Demonstrated

This project demonstrates experience with:

- Python Development
- Flask Web Applications
- Authentication & Session Management
- Workflow Automation
- Business Process Digitization
- Data Validation
- Excel Automation
- Dashboard Development
- Multi-user Application Design
- Concurrent File Access
- Secure Configuration Management
- Software Architecture

---

## 🔮 Future Improvements

- Database backend (SQLite/PostgreSQL)
- Role-based access control
- Individual user accounts
- REST API
- Audit logging
- Report generation
- Docker deployment
- Unit and integration testing

---

## 📄 License

This project is licensed under the **MIT License**.

---

## ⚠️ Disclaimer

This repository has been sanitized for public portfolio purposes. Any organization-specific information, operational data, credentials, infrastructure configuration, and sensitive deployment assets have been removed or generalized.

The repository is intended solely to demonstrate software engineering, application development, and workflow automation skills.
