import sqlite3
import pandas as pd
import os

EXCEL_PATH = r"C:\Users\Admin\Downloads\pharmacy_app\Ibukunolu new pharmacy database.xlsx"
DB_PATH = "pharmacy.db"

# ---------------------------------------------------------
# DELETE OLD DB
# ---------------------------------------------------------
if os.path.exists(DB_PATH):
    os.remove(DB_PATH)
    print("Old pharmacy.db deleted.")

# ---------------------------------------------------------
# CREATE NEW DB + TABLES
# ---------------------------------------------------------
conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

cur.executescript("""
CREATE TABLE IF NOT EXISTS DrugList (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    drug_name TEXT,
    sales_price REAL,
    starting_balance REAL,
    categories TEXT,
    cost_price REAL
);

CREATE TABLE IF NOT EXISTS Stock (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT,
    time TEXT,
    drug_name TEXT,
    categories TEXT,
    opening_balance REAL,
    transaction_type TEXT,
    quantity_sold REAL,
    quantity_remaining REAL,
    cost_price REAL,
    sales_price REAL,
    customer_type TEXT,
    transaction_id TEXT,
    remark TEXT,
    bill REAL,
    sales_price_sum REAL,
    sum_cost_value REAL
);

CREATE TABLE IF NOT EXISTS InpatientData (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT,
    drug_name TEXT,
    patient_name TEXT,
    categories TEXT,
    sales_price REAL,
    transaction_id TEXT,
    bill_total REAL,
    qty_sold REAL
);

CREATE TABLE IF NOT EXISTS DutyLog (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT,
    login_time TEXT,
    logout_time TEXT
);

CREATE TABLE IF NOT EXISTS Users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    [User Id] TEXT UNIQUE,
    Password TEXT,
    Name TEXT,
    Role TEXT,
    Status TEXT DEFAULT 'Active'
);

CREATE TABLE IF NOT EXISTS Bill (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    date TEXT,
    drug_name TEXT,
    patient_name TEXT,
    categories TEXT,
    sales_price REAL,
    transaction_id TEXT,
    bill_total REAL,
    source TEXT
);
""")

conn.commit()
print("New tables created.")

# ---------------------------------------------------------
# IMPORT EXCEL SHEETS
# ---------------------------------------------------------
def load_sheet(name):
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=name)
        df.columns = df.columns.astype(str).str.strip()
        return df
    except:
        print(f"Sheet '{name}' not found.")
        return pd.DataFrame()

# ------------------ DrugList -----------------------------
df = load_sheet("Drug List")
if not df.empty:
    df = df.rename(columns={
        "Drug Name": "drug_name",
        "Sales Price": "sales_price",
        "Starting Balance": "starting_balance",
        "Categories": "categories",
        "Cost Price": "cost_price"
    })
    df.to_sql("DrugList", conn, if_exists="append", index=False)
    print("DrugList imported.")

# ------------------ StockView -----------------------------
df = load_sheet("StockView")
if not df.empty:
    df = df.rename(columns={
        "DATE": "date",
        "TIME": "time",
        "DRUG NAME": "drug_name",
        "CATEGORIES": "categories",
        "OPENING BALANCE": "opening_balance",
        "TRANSACTION TYPE": "transaction_type",
        "QUANTITY SOLD": "quantity_sold",
        "QUANTITY REMAINING": "quantity_remaining",
        "COST PRICE": "cost_price",
        "SALES PRICE": "sales_price",
        "CUSTOMER TYPE": "customer_type",
        "TRANSACTION ID": "transaction_id",
        "REMARK": "remark",
        "BILL": "bill",
        "Sum of Sales Value": "sales_price_sum",
        "Sum of Cost Value": "sum_cost_value"
    })
    df.to_sql("Stock", conn, if_exists="append", index=False)
    print("Stock imported.")

# ------------------ InpatientData ------------------------
df = load_sheet("Inpatient Data")
if not df.empty:
    df = df.rename(columns={
        "Date": "date",
        "Drug Name": "drug_name",
        "Patient Name": "patient_name",
        "Categories": "categories",
        "Sales Price": "sales_price",
        "Transaction ID": "transaction_id",
        "Bill Total": "bill_total",
        "Qty Sold": "qty_sold"
    })
    df.to_sql("InpatientData", conn, if_exists="append", index=False)
    print("InpatientData imported.")

# ------------------ DutyLog ------------------------------
df = load_sheet("DutyLog")
if not df.empty:
    df = df.rename(columns={
        "User Id": "user_id",
        "Login Time": "login_time",
        "Logout Time": "logout_time"
    })
    df.to_sql("DutyLog", conn, if_exists="append", index=False)
    print("DutyLog imported.")

# ------------------ Users ------------------------------
import bcrypt

df = load_sheet("Users Data")
if not df.empty:
    df = df.rename(columns={
        "User Id": "User Id",
        "Password": "Password",
        "Name": "Name",
        "Role": "Role"
    })

    # Hash passwords
    hashed_rows = []
    for _, row in df.iterrows():
        user_id = str(row["User Id"]).strip()
        name = str(row["Name"]).strip()
        role = str(row["Role"]).strip()
        raw_password = str(row["Password"]).strip()

        hashed_pw = bcrypt.hashpw(raw_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

        hashed_rows.append((user_id, hashed_pw, name, role, "Active"))

    cur.executemany("""
        INSERT INTO Users ([User Id], Password, Name, Role, Status)
        VALUES (?, ?, ?, ?, ?)
    """, hashed_rows)

    conn.commit()
    print("Users imported with hashed passwords.")

# ------------------ Bill ------------------------------
df = load_sheet("Bill")
if not df.empty:
    df = df.rename(columns={
        "Date": "date",
        "Drug Name": "drug_name",
        "Patient Name": "patient_name",
        "Categories": "categories",
        "Sales Price": "sales_price",
        "Transaction ID": "transaction_id",
        "Bill Total": "bill_total",
        "Source": "source"
    })
    df.to_sql("Bill", conn, if_exists="append", index=False)
    print("Bill imported.")

conn.close()
print("pharmacy.db rebuild complete.")
