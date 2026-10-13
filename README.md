<div align="center">

# 🏪 IntersportApp

### Customer orders, store follow-ups, and email reminders in one place.

A Windows desktop application for Intersport and Sportstore teams, built with **Python**, **PySide6**, and **SQLite**.

**Version 1.1.3.0 · Dutch interface · Local data storage**

</div>

---

## Overview

IntersportApp helps store staff manage customer orders and requests from registration through collection. Record customer and product details, track payment and contact status, and keep outstanding orders visible in a searchable dashboard.

Customer confirmation emails and scheduled reminder summaries support day-to-day follow-up, while switchable Intersport and Sportstore branding keeps the interface consistent with the store.

## ✨ Features

| Feature | Description |
| --- | --- |
| **Order registration** | Capture customer name, salesperson, contact details, order description, size, article number, EAN, and recommended retail price. |
| **Store coordination** | Record the contacted branch or supplier and the colleague involved. |
| **Status tracking** | Use color-coded contact, payment, and order statuses. |
| **Searchable dashboard** | Search across order details and sort tables by column. |
| **Completion tracking** | View open and completed orders separately, with options to complete or reopen an order. |
| **Customer emails** | Automatically attempt a confirmation email when saving an order with an email address; send an arrival notification through the dashboard. |
| **Reminder summaries** | Combine eligible outstanding orders into a single email to the configured store address. |
| **Brand themes** | Switch between Intersport and Sportstore styling and logos. |
| **Local persistence** | Store orders in SQLite and settings in JSON. |
| **Recovery support** | Migrate data from recognized older locations and create dated database backups at startup. |
| **Diagnostics** | Record application activity and crash information in persistent log files. |

## 🚀 Getting Started

### Run the Windows application

If a packaged build is available, open:

```text
dist/IntersportApp/IntersportApp.exe
```

Keep the executable and its accompanying `_internal` folder together. The packaged application includes the Python runtime and required dependencies.

### Run from source

Use Python **3.10 or newer**. The existing packaged build includes Python 3.11.

Run the following commands in PowerShell from the project directory:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install PySide6
.\.venv\Scripts\python.exe app.py
```

The database is initialized automatically on first launch.

## 🧭 Using the Application

### 1. Configure the store

Open **Instellingen** to:

- Set **Email voor reminders**, the address that receives store reminder summaries.
- Choose the **Intersport** or **Sportstore** theme.
- Save your settings.
- Send a test email to verify email delivery.
- View or open the active data folder.

### 2. Register an order

Open **Invoeren** and complete the required fields:

- Customer name
- Salesperson
- Email address or phone number
- Order description
- Article number
- EAN code

Optionally add the size, recommended retail price, branch, colleague, and relevant statuses.

Click **Opslaan** to save the order and open the dashboard. If the customer has an email address, the application attempts to send a confirmation email.

An email failure does not undo the saved order.

### 3. Manage orders

Open **Dashboard** to review **Openstaande bestellingen** and **Afgeronde bestellingen**.

- Use the search field to find an order.
- Click a column header to sort.
- Double-click an order to edit it.
- Right-click an open order to edit it, mark it complete, or send a customer arrival email.
- Right-click a completed order and choose **Niet afgerond** to reopen it.
- Click **Ververs** to reload the dashboard.

> Phone numbers can be stored and validated. Automated WhatsApp and SMS delivery are not currently implemented.

## 📬 Email and Reminders

### Customer messages

The application supports two customer email workflows:

| Message | Trigger |
| --- | --- |
| **Order confirmation** | Automatically attempted when a new order is saved with a valid email address. |
| **Arrival notification** | Sent through the dashboard action **Stuur automatisch bericht naar klant**. |

### Store reminders

Reminder summaries are sent to the address configured in **Instellingen**.

An order qualifies when:

- It remains open.
- It was created at least **7 days** ago.
- No reminder has been sent, or the previous reminder was sent at least **7 days** ago.

The application checks at startup and then approximately every hour while it remains open. **Verstuur reminders nu** runs the same eligibility check manually.

Eligible orders are combined into one summary email. If a Gmail daily sending limit is detected, email sending enters a **24-hour cooldown**.

> Automatic reminders require the application to be running. Email delivery requires an internet connection and valid SMTP credentials.

### SMTP configuration

The email configuration is defined near the top of `app.py`:

| Setting | Purpose |
| --- | --- |
| `SMTP_SERVER` | SMTP host; currently `smtp.gmail.com`. |
| `SMTP_PORT` | SSL port; currently `465`. |
| `SENDER_EMAIL` | Sender account. |
| `APP_PASSWORD` | SMTP authentication credential. |

**The current source assigns `APP_PASSWORD` directly.** To configure it through an environment variable, replace that assignment with:

```python
APP_PASSWORD = os.getenv("EMAIL_PASS", "").strip()
```

Then set the variable before launching the application:

```powershell
$env:EMAIL_PASS = "your-app-password"
.\.venv\Scripts\python.exe app.py
```

Keep credentials out of source control and documentation.

## 💾 Data Storage and Backups

Application data is stored separately from the executable.

| Installation type | Default data location |
| --- | --- |
| **Standard Windows execution** | `%LOCALAPPDATA%\IntersportApp` |
| **MSIX-packaged execution** | `%LOCALAPPDATA%\Packages\<package-family>\LocalState\IntersportApp` |

The active folder is visible under **Instellingen → Actieve gegevensmap**.

```text
IntersportApp/
├── store.db                 # Customer orders
├── settings.json            # Application settings
├── backups/
│   └── store-YYYYMMDD.db     # Dated database backups
└── logs/
    ├── IntersportApp.log
    └── IntersportApp-crash.log
```

### Backup behavior

At startup, the application creates a database backup if one does not already exist for that date. It retains the **seven most recent backup files**.

Recognized legacy database and settings locations are checked automatically when the active data files do not yet exist.

### Custom data directory

Set `INTERSPORT_DATA_DIR` before launch to use another location:

```powershell
$env:INTERSPORT_DATA_DIR = "C:\IntersportAppData"
.\.venv\Scripts\python.exe app.py
```

The application uses local file storage. It does not provide a central server or built-in synchronization between stores.

## 🛠️ Development

### Technology

| Component | Technology |
| --- | --- |
| Application | Python |
| Desktop interface | PySide6 / Qt Widgets |
| Database | SQLite |
| Settings | JSON |
| Email | Python `smtplib` with SMTP over SSL |
| Styling | Qt Style Sheets (`.qss`) |
| Windows distribution | PyInstaller |

### Project structure

```text
IntersportApp/
├── app.py                         # Application logic and desktop interface
├── test_storage.py                # Storage, email, and interface tests
├── IntersportApp.spec              # PyInstaller build configuration
├── intersport.qss                  # Intersport theme
├── Sportstore.qss                  # Sportstore theme
├── KOBALT ZONDER PLAATSNAAM.png     # Intersport logo
├── Sportstore kobalt.png           # Sportstore logo
├── build/                         # Generated build files
└── dist/                          # Packaged application
```

### Run tests

```powershell
.\.venv\Scripts\python.exe -m unittest test_storage -v
```

The suite covers legacy data migration, database writes and locking, order-form saving, reminder digests and scheduling, MSIX storage paths, and email quota cooldown behavior.

Tests use temporary data directories, mocked email delivery, and Qt's offscreen platform.

### Build the Windows application

Install PyInstaller and build using the included specification:

```powershell
.\.venv\Scripts\python.exe -m pip install pyinstaller
.\.venv\Scripts\python.exe -m PyInstaller --noconfirm IntersportApp.spec
```

The resulting executable is created at:

```text
dist/IntersportApp/IntersportApp.exe
```

The specification bundles both themes and logo assets and produces a windowed application without a console.

## 🔎 Troubleshooting

| Issue | What to check |
| --- | --- |
| **Email is not delivered** | Verify SMTP credentials, internet access, and the test-email result in **Instellingen**. |
| **Reminders are not sent** | Confirm a reminder address is saved, the app is running, and orders meet the seven-day eligibility rules. |
| **Email sending is paused** | Check whether the application reported a daily sending limit and cooldown. |
| **An order cannot be saved** | Review the error message and check whether another process is locking the database. |
| **Previous orders appear missing** | Check the active data folder and any `INTERSPORT_DATA_DIR` override. |
| **The app reports it is already open** | Close the existing instance before opening another against the same data directory. |
| **The app fails to start** | Inspect `logs/IntersportApp.log` and `logs/IntersportApp-crash.log` in the data folder. |

If the usual diagnostics folder is unavailable, the application attempts to write logs under `%TEMP%\IntersportApp\logs`.

## 👤 Author

**Alexander Brinkman**

© 2026 Alexander Brinkman. All rights reserved.
