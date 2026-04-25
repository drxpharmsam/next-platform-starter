"""
Medicine Database Autobot
=========================
Reads medicine data from an Excel spreadsheet and uploads it to a SQL database.
Duplicate medicine names are automatically skipped.

Usage:
    python medicine_autobot.py                          # uses config.ini settings
    python medicine_autobot.py --file medicines.xlsx    # override Excel file path
    python medicine_autobot.py --help                   # show all options
"""

import argparse
import configparser
import logging
import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from sqlalchemy import (
    Column,
    DateTime,
    Float,
    Integer,
    String,
    Text,
    create_engine,
    inspect,
    text,
)
from sqlalchemy.exc import SQLAlchemyError
from sqlalchemy.orm import DeclarativeBase, Session

# ---------------------------------------------------------------------------
# Logging setup
# ---------------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("autobot.log", encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# ORM model
# ---------------------------------------------------------------------------


class Base(DeclarativeBase):
    pass


class Medicine(Base):
    """Maps to the 'medicines' table in the database."""

    __tablename__ = "medicines"

    id = Column(Integer, primary_key=True, autoincrement=True)
    name = Column(String(255), nullable=False, unique=True)
    generic_name = Column(String(255))
    brand_name = Column(String(255))
    manufacturer = Column(String(255))
    category = Column(String(100))
    dosage_form = Column(String(100))       # tablet, capsule, syrup, injection …
    strength = Column(String(100))          # e.g. "500 mg", "10 mg/5 mL"
    unit_price = Column(Float)
    stock_quantity = Column(Integer)
    expiry_date = Column(String(50))        # stored as text for flexibility
    description = Column(Text)
    side_effects = Column(Text)
    contraindications = Column(Text)
    storage_conditions = Column(String(255))
    created_at = Column(DateTime, default=datetime.utcnow)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


# ---------------------------------------------------------------------------
# Required columns that MUST be present in the spreadsheet
# ---------------------------------------------------------------------------
REQUIRED_COLUMNS = {"name"}

# Mapping from spreadsheet column headers (case-insensitive) → ORM attribute names.
# Add or modify entries here to match your actual spreadsheet headers.
COLUMN_MAP = {
    "name": "name",
    "medicine name": "name",
    "drug name": "name",
    "generic name": "generic_name",
    "generic": "generic_name",
    "brand name": "brand_name",
    "brand": "brand_name",
    "manufacturer": "manufacturer",
    "company": "manufacturer",
    "category": "category",
    "drug category": "category",
    "dosage form": "dosage_form",
    "form": "dosage_form",
    "strength": "strength",
    "dose": "strength",
    "unit price": "unit_price",
    "price": "unit_price",
    "cost": "unit_price",
    "stock": "stock_quantity",
    "stock quantity": "stock_quantity",
    "quantity": "stock_quantity",
    "expiry date": "expiry_date",
    "expiry": "expiry_date",
    "exp date": "expiry_date",
    "description": "description",
    "details": "description",
    "side effects": "side_effects",
    "contraindications": "contraindications",
    "storage": "storage_conditions",
    "storage conditions": "storage_conditions",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def load_config(config_path: str = "config.ini") -> configparser.ConfigParser:
    cfg = configparser.ConfigParser()
    if os.path.exists(config_path):
        cfg.read(config_path, encoding="utf-8")
        log.info("Loaded configuration from %s", config_path)
    else:
        log.warning("Config file '%s' not found – using built-in defaults.", config_path)
    return cfg


def build_connection_string(cfg: configparser.ConfigParser) -> str:
    """Return a SQLAlchemy connection URL from config or defaults to SQLite."""
    db_type = cfg.get("database", "type", fallback="sqlite").lower()

    if db_type == "sqlite":
        db_path = cfg.get("database", "path", fallback="medicines.db")
        return f"sqlite:///{db_path}"

    if db_type in ("mysql", "mariadb"):
        driver = cfg.get("database", "driver", fallback="pymysql")
        user = cfg.get("database", "user", fallback="root")
        password = cfg.get("database", "password", fallback="")
        host = cfg.get("database", "host", fallback="localhost")
        port = cfg.get("database", "port", fallback="3306")
        dbname = cfg.get("database", "dbname", fallback="pharmacy")
        return f"mysql+{driver}://{user}:{password}@{host}:{port}/{dbname}"

    if db_type == "postgresql":
        driver = cfg.get("database", "driver", fallback="psycopg2")
        user = cfg.get("database", "user", fallback="postgres")
        password = cfg.get("database", "password", fallback="")
        host = cfg.get("database", "host", fallback="localhost")
        port = cfg.get("database", "port", fallback="5432")
        dbname = cfg.get("database", "dbname", fallback="pharmacy")
        return f"postgresql+{driver}://{user}:{password}@{host}:{port}/{dbname}"

    if db_type == "mssql":
        driver = cfg.get("database", "driver", fallback="pyodbc")
        user = cfg.get("database", "user", fallback="sa")
        password = cfg.get("database", "password", fallback="")
        host = cfg.get("database", "host", fallback="localhost")
        port = cfg.get("database", "port", fallback="1433")
        dbname = cfg.get("database", "dbname", fallback="pharmacy")
        return (
            f"mssql+{driver}://{user}:{password}@{host}:{port}/{dbname}"
            "?driver=ODBC+Driver+17+for+SQL+Server"
        )

    raise ValueError(f"Unsupported database type: '{db_type}'")


def read_excel(file_path: str, sheet_name) -> pd.DataFrame:
    """Load the spreadsheet and normalise column headers."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Excel file not found: {file_path}")

    log.info("Reading spreadsheet: %s  (sheet: %s)", file_path, sheet_name)
    df = pd.read_excel(path, sheet_name=sheet_name, dtype=str)

    # Normalise column names: strip whitespace, lowercase
    df.columns = [str(c).strip().lower() for c in df.columns]

    # Drop completely empty rows
    df.dropna(how="all", inplace=True)

    log.info("Spreadsheet loaded – %d row(s) found.", len(df))
    return df


def map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Rename spreadsheet columns to ORM attribute names using COLUMN_MAP."""
    rename = {}
    for col in df.columns:
        mapped = COLUMN_MAP.get(col)
        if mapped:
            rename[col] = mapped

    df = df.rename(columns=rename)

    # Check required columns exist after mapping
    missing = REQUIRED_COLUMNS - set(df.columns)
    if missing:
        raise ValueError(
            f"Required column(s) missing from spreadsheet: {missing}\n"
            f"Available columns after mapping: {list(df.columns)}\n"
            "Tip: Ensure your spreadsheet has a column named 'name' or 'Medicine Name'."
        )

    return df


def clean_value(value):
    """Convert NaN / empty strings to None; strip whitespace from strings."""
    if pd.isna(value):
        return None
    v = str(value).strip()
    return None if v in ("", "nan", "None", "NaN") else v


def upload_medicines(df: pd.DataFrame, session: Session) -> dict:
    """Insert rows that are not already in the database. Returns stats."""
    stats = {"inserted": 0, "skipped_duplicate": 0, "skipped_no_name": 0, "errors": 0}

    # Pre-fetch all existing medicine names (lower-cased) for fast duplicate check
    existing_names = {
        row[0].strip().lower()
        for row in session.execute(text("SELECT name FROM medicines")).fetchall()
    }
    log.info("Existing medicines in DB: %d", len(existing_names))

    orm_columns = {c.key for c in Medicine.__table__.columns} - {"id", "created_at", "updated_at"}

    for idx, row in df.iterrows():
        row_num = idx + 2  # Excel row number (1-indexed header + 1)

        name_raw = clean_value(row.get("name", ""))
        if not name_raw:
            log.warning("Row %d: skipped – 'name' column is empty.", row_num)
            stats["skipped_no_name"] += 1
            continue

        if name_raw.lower() in existing_names:
            log.info("Row %d: skipped duplicate – '%s'", row_num, name_raw)
            stats["skipped_duplicate"] += 1
            continue

        # Build ORM instance from available columns
        kwargs = {"name": name_raw}
        for col in orm_columns - {"name"}:
            if col in df.columns:
                kwargs[col] = clean_value(row.get(col))

        # Handle numeric conversions
        if kwargs.get("unit_price") is not None:
            try:
                kwargs["unit_price"] = float(kwargs["unit_price"])
            except (ValueError, TypeError):
                log.warning("Row %d: invalid unit_price value '%s' – set to NULL.", row_num, kwargs["unit_price"])
                kwargs["unit_price"] = None

        if kwargs.get("stock_quantity") is not None:
            try:
                kwargs["stock_quantity"] = int(float(kwargs["stock_quantity"]))
            except (ValueError, TypeError):
                log.warning("Row %d: invalid stock_quantity value '%s' – set to NULL.", row_num, kwargs["stock_quantity"])
                kwargs["stock_quantity"] = None

        try:
            medicine = Medicine(**kwargs)
            session.add(medicine)
            session.flush()  # catch constraint violations early
            existing_names.add(name_raw.lower())
            log.info("Row %d: inserted – '%s'", row_num, name_raw)
            stats["inserted"] += 1
        except SQLAlchemyError as exc:
            session.rollback()
            log.error("Row %d: DB error for '%s' – %s", row_num, name_raw, exc)
            stats["errors"] += 1

    return stats


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------


def main():
    parser = argparse.ArgumentParser(
        description="Medicine Database Autobot – uploads Excel data to SQL, skipping duplicates."
    )
    parser.add_argument("--file", "-f", help="Path to Excel file (overrides config.ini)")
    parser.add_argument("--sheet", "-s", default=None, help="Sheet name or index (default: first sheet)")
    parser.add_argument("--config", "-c", default="config.ini", help="Path to config file (default: config.ini)")
    parser.add_argument("--db", "-d", help="SQLAlchemy connection URL (overrides config.ini)")
    parser.add_argument("--dry-run", action="store_true", help="Parse and validate without writing to DB")
    args = parser.parse_args()

    # --- Load configuration ---
    cfg = load_config(args.config)

    excel_file = args.file or cfg.get("excel", "file_path", fallback="medicines.xlsx")
    sheet_name = args.sheet or cfg.get("excel", "sheet_name", fallback=0)

    # Sheet name can be an integer index
    try:
        sheet_name = int(sheet_name)
    except (ValueError, TypeError):
        pass

    connection_url = args.db or build_connection_string(cfg)

    # --- Read & validate spreadsheet ---
    try:
        df = read_excel(excel_file, sheet_name)
        df = map_columns(df)
    except (FileNotFoundError, ValueError) as exc:
        log.error("%s", exc)
        sys.exit(1)

    if args.dry_run:
        log.info("DRY RUN – no data will be written to the database.")
        log.info("Columns detected: %s", list(df.columns))
        log.info("First 5 rows:\n%s", df.head().to_string())
        sys.exit(0)

    # --- Connect to database & upload ---
    log.info("Connecting to database …")
    try:
        engine = create_engine(connection_url, echo=False)
        Base.metadata.create_all(engine)  # create table if it doesn't exist
        log.info("Table 'medicines' ready.")
    except SQLAlchemyError as exc:
        log.error("Failed to connect / create table: %s", exc)
        sys.exit(1)

    with Session(engine) as session:
        stats = upload_medicines(df, session)
        session.commit()

    # --- Summary ---
    print("\n" + "=" * 50)
    print("  AUTOBOT UPLOAD SUMMARY")
    print("=" * 50)
    print(f"  ✅  Inserted       : {stats['inserted']}")
    print(f"  ⚠️   Duplicates skipped : {stats['skipped_duplicate']}")
    print(f"  ⛔  Skipped (no name) : {stats['skipped_no_name']}")
    print(f"  ❌  Errors         : {stats['errors']}")
    print("=" * 50)
    print(f"  Log saved to: autobot.log")
    print("=" * 50 + "\n")

    if stats["errors"] > 0:
        sys.exit(2)


if __name__ == "__main__":
    main()
