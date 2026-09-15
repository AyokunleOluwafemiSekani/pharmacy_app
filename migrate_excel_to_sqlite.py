import pandas as pd
import sqlite3

EXCEL_FILE = r"C:\Users\Admin\Downloads\pharmacy_app\Ibukunolu new pharmacy database.xlsx"
DB_FILE = "pharmacy.db"

print("Loading Excel sheets...")

# Load sheets
drug_df = pd.read_excel(EXCEL_FILE, sheet_name="Drug List")
stock_df = pd.read_excel(EXCEL_FILE, sheet_name="StockView")
inpatient_df = pd.read_excel(EXCEL_FILE, sheet_name="Inpatient Data")
bill_df = pd.read_excel(EXCEL_FILE, sheet_name="Bill")
duty_df = pd.read_excel(EXCEL_FILE, sheet_name="DutyLog")

# Clean column names
drug_df.columns = drug_df.columns.astype(str).str.strip()
stock_df.columns = stock_df.columns.astype(str).str.strip()
inpatient_df.columns = inpatient_df.columns.astype(str).str.strip()
bill_df.columns = bill_df.columns.astype(str).str.strip()
duty_df.columns = duty_df.columns.astype(str).str.strip()

# Convert all date/time columns to strings
if "DATE" in stock_df.columns:
    stock_df["DATE"] = stock_df["DATE"].astype(str)
if "TIME" in stock_df.columns:
    stock_df["TIME"] = stock_df["TIME"].astype(str)

if "Date" in inpatient_df.columns:
    inpatient_df["Date"] = inpatient_df["Date"].astype(str)

if "Date" in bill_df.columns:
    bill_df["Date"] = bill_df["Date"].astype(str)

if "Login Time" in duty_df.columns:
    duty_df["Login Time"] = duty_df["Login Time"].astype(str)
if "Logout Time" in duty_df.columns:
    duty_df["Logout Time"] = duty_df["Logout Time"].astype(str)

# Connect to SQLite
conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

# -----------------------------
# MIGRATE DRUG LIST
# -----------------------------
print("Migrating Drug List...")
for _, row in drug_df.iterrows():
    cursor.execute("""
        INSERT INTO DrugList (
            drug_name, sales_price, starting_balance,
            categories, cost_price
        )
        VALUES (?, ?, ?, ?, ?)
    """, (
        row.get("Drug Name"),
        row.get("Sales Price"),
        row.get("Starting Balance"),
        row.get("Categories"),
        row.get("Cost Price")
    ))

# -----------------------------
# MIGRATE STOCKVIEW → STOCK
# -----------------------------
print("Migrating StockView → Stock...")
for _, row in stock_df.iterrows():
    cursor.execute("""
        INSERT INTO Stock (
            date, time, drug_name, categories,
            opening_balance, transaction_type,
            quantity_sold, quantity_remaining,
            cost_price, sales_price, customer_type,
            transaction_id, remark, bill,
            sum_sales_value, sum_cost_value
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("DATE"),
        row.get("TIME"),
        row.get("DRUG NAME"),
        row.get("CATEGORIES"),
        row.get("OPENING BALANCE"),
        row.get("TRANSACTION TYPE"),
        row.get("QUANTITY SOLD"),
        row.get("QUANTITY REMAINING"),
        row.get("COST PRICE"),
        row.get("SALES PRICE"),
        row.get("CUSTOMER TYPE"),
        row.get("TRANSACTION ID"),
        row.get("REMARK"),
        row.get("BILL"),
        row.get("Sum of Sales Value"),
        row.get("Sum of Cost Value")
    ))

# -----------------------------
# MIGRATE INPATIENT DATA
# -----------------------------
print("Migrating Inpatient Data...")
for _, row in inpatient_df.iterrows():
    cursor.execute("""
        INSERT INTO InpatientData (
            date, drug_name, patient_name, categories,
            sales_price, transaction_id, bill_total, qty_sold
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("Date"),
        row.get("Drug Name"),
        row.get("Patient Name"),
        row.get("Categories"),
        row.get("Sales Price"),
        row.get("Transaction ID"),
        row.get("Bill Total"),
        row.get("Qty Sold")
    ))

# -----------------------------
# MIGRATE BILL
# -----------------------------
print("Migrating Bill...")
for _, row in bill_df.iterrows():
    cursor.execute("""
        INSERT INTO Bill (
            date, drug_name, patient_name, categories,
            sales_price, transaction_id, bill_total, source
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        row.get("Date"),
        row.get("Drug Name"),
        row.get("Patient Name"),
        row.get("Categories"),
        row.get("Sales Price"),
        row.get("Transaction ID"),
        row.get("Bill Total"),
        row.get("Source")
    ))

# -----------------------------
# MIGRATE DUTY LOG
# -----------------------------
print("Migrating DutyLog...")
for _, row in duty_df.iterrows():
    cursor.execute("""
        INSERT INTO DutyLog (user_id, login_time, logout_time)
        VALUES (?, ?, ?)
    """, (
        row.get("User Id"),
        row.get("Login Time"),
        row.get("Logout Time")
    ))

conn.commit()
conn.close()

print("\nMigration completed successfully!")
print("All Excel sheets have been imported into SQLite.")
