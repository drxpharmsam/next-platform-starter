# Medicine Database Autobot

A Python script that reads medicine data from an Excel spreadsheet and uploads it to a SQL database automatically — skipping any duplicate entries.

---

## Features

- 📋 Reads `.xlsx` / `.xls` spreadsheets
- 🗄️ Supports **SQLite** (no server needed), **MySQL/MariaDB**, **PostgreSQL**, and **MS SQL Server**
- 🚫 **Skips duplicate medicine names** automatically
- 📝 Full log written to `autobot.log`
- 🔁 Safe to re-run — already-uploaded rows are never duplicated
- 🖥️ Runs directly inside **VS Code** terminal

---

## File Structure

```
📁 your-project/
├── medicine_autobot.py      ← main script
├── create_sample_excel.py   ← helper: generates a sample Excel file
├── config.ini               ← database & file configuration
├── requirements.txt         ← Python dependencies
├── medicines.xlsx           ← YOUR spreadsheet (you create/edit this)
├── medicines.db             ← SQLite database (auto-created)
└── autobot.log              ← log file (auto-created)
```

---

## Quick Start (VS Code)

### 1 — Install Python

Make sure Python 3.9+ is installed. In the VS Code terminal:
```bash
python --version
```

### 2 — Install dependencies

Open the VS Code terminal (`Ctrl + ~`) and run:
```bash
pip install -r requirements.txt
```

If you want MySQL support, also run:
```bash
pip install PyMySQL
```

If you want PostgreSQL support:
```bash
pip install psycopg2-binary
```

### 3 — Generate a sample Excel template (optional)

```bash
python create_sample_excel.py
```
This creates `medicines.xlsx` with example data and all the correct column headers. Open it in Excel, delete the sample rows, and fill in your real medicines.

### 4 — Configure the database

Edit `config.ini`:

**For SQLite (simplest — no server needed):**
```ini
[database]
type = sqlite
path = medicines.db
```

**For MySQL:**
```ini
[database]
type     = mysql
driver   = pymysql
host     = localhost
port     = 3306
user     = root
password = your_password
dbname   = pharmacy
```

**For PostgreSQL:**
```ini
[database]
type     = postgresql
driver   = psycopg2
host     = localhost
port     = 5432
user     = postgres
password = your_password
dbname   = pharmacy
```

### 5 — Run the autobot

```bash
python medicine_autobot.py
```

That's it! The script will:
1. Read your Excel file
2. Connect to the database (create the `medicines` table if it doesn't exist)
3. Insert new rows, skipping any name that already exists
4. Print a summary

---

## Excel Spreadsheet Format

Your spreadsheet must have **at minimum** a column named `Name` (or `Medicine Name` / `Drug Name`). All other columns are optional.

| Column Header | Notes |
|---|---|
| **Name** *(required)* | Unique medicine name — duplicates are skipped |
| Generic Name | INN / generic drug name |
| Brand Name | Trade / brand name |
| Manufacturer | Pharmaceutical company |
| Category | e.g. Antibiotic, Analgesic |
| Dosage Form | Tablet, Capsule, Syrup, Injection … |
| Strength | e.g. `500 mg`, `10 mg/5 mL` |
| Unit Price | Numeric price per unit |
| Stock Quantity | Integer quantity in stock |
| Expiry Date | e.g. `2027-06-30` |
| Description | Free text description |
| Side Effects | Free text |
| Contraindications | Free text |
| Storage Conditions | e.g. `Store below 25°C` |

> Column headers are **case-insensitive** and leading/trailing spaces are trimmed automatically.

---

## Command-line Options

```
python medicine_autobot.py [options]

Options:
  -f, --file FILE     Path to Excel file (overrides config.ini)
  -s, --sheet SHEET   Sheet name or index, e.g. "Sheet1" or 0  (default: first sheet)
  -c, --config FILE   Path to config.ini  (default: config.ini)
  -d, --db URL        SQLAlchemy URL, e.g. sqlite:///my.db  (overrides config.ini)
  --dry-run           Parse and validate without writing to the database
  -h, --help          Show this help message
```

### Examples

```bash
# Use a different Excel file
python medicine_autobot.py --file /path/to/my_medicines.xlsx

# Use a named sheet
python medicine_autobot.py --sheet "Medicines Jan 2025"

# Validate without uploading
python medicine_autobot.py --dry-run

# Connect directly to a MySQL database without editing config.ini
python medicine_autobot.py --db "mysql+pymysql://root:pass@localhost/pharmacy"
```

---

## Database Table Schema

The script automatically creates a table called `medicines`:

| Column | Type | Notes |
|---|---|---|
| id | INTEGER | Auto-increment primary key |
| name | VARCHAR(255) | **Unique**, required |
| generic_name | VARCHAR(255) | |
| brand_name | VARCHAR(255) | |
| manufacturer | VARCHAR(255) | |
| category | VARCHAR(100) | |
| dosage_form | VARCHAR(100) | |
| strength | VARCHAR(100) | |
| unit_price | FLOAT | |
| stock_quantity | INTEGER | |
| expiry_date | VARCHAR(50) | |
| description | TEXT | |
| side_effects | TEXT | |
| contraindications | TEXT | |
| storage_conditions | VARCHAR(255) | |
| created_at | DATETIME | Auto-set on insert |
| updated_at | DATETIME | Auto-set on update |

---

## Re-running Safely

You can run the script as many times as you like. Any medicine whose `name` already exists in the database will be **skipped** — it will never be inserted twice.

This means you can:
- Add new rows to your spreadsheet and re-run — only the new ones are inserted
- Fix typos in other columns and re-run — existing rows are not touched
- Schedule the script to run automatically (see below)

---

## Scheduling (optional)

**Windows Task Scheduler:**
1. Open Task Scheduler → Create Basic Task
2. Set trigger (e.g. daily)
3. Action: `python C:\path\to\medicine_autobot.py`

**Linux/macOS cron:**
```bash
# Run every day at 8 AM
0 8 * * * /usr/bin/python3 /path/to/medicine_autobot.py >> /var/log/autobot.log 2>&1
```

---

## Troubleshooting

| Problem | Solution |
|---|---|
| `ModuleNotFoundError: No module named 'pandas'` | Run `pip install -r requirements.txt` |
| `FileNotFoundError: medicines.xlsx` | Check the `file_path` in `config.ini` or pass `--file` |
| `Required column(s) missing: {'name'}` | Add a column named **Name** or **Medicine Name** to your spreadsheet |
| MySQL connection refused | Check host, port, user, password in `config.ini` and that MySQL is running |
| `pip` not found | Use `python -m pip install -r requirements.txt` |
