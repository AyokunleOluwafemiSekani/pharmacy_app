import os
import bcrypt
from flask import Flask, render_template, request, redirect, session
import pandas as pd
import sqlite3
from datetime import datetime
from backup_manager import create_backup
from auto_backup_manager import load_settings, save_settings, run_auto_backup_if_due
from threading import Thread
import time
from kpi_manager import get_kpis


app = Flask(__name__)
app.secret_key = "super_secret_key"

# Excel file (ONLY used for initial migration)
EXCEL_PATH = r"C:\Users\Admin\Downloads\pharmacy_app\Ibukunolu new pharmacy database.xlsx"
DB_PATH = "pharmacy.db"
# ---------------------------------------------------------
# ROLE PERMISSIONS
# ---------------------------------------------------------
ROLE_PERMISSIONS = {
    "Admin": {
        "pos": True,
        "stock": True,
        "dutylog": True,
        "admin_panel": True,
        "user_management": True,
        "inpatient": True,
        "kpi": True,
        "drug_list": True,
        "drug_entry": True
    },

    "Administrator": {
        "pos": True,
        "stock": True,
        "dutylog": True,
        "admin_panel": True,
        "user_management": True,
        "inpatient": True,
        "kpi": True,
        "drug_list": True,
        "drug_entry": True
    },

    "Auditor": {
        "pos": True,
        "stock": True,
        "dutylog": False,
        "admin_panel": False,
        "user_management": False,
        "inpatient": True,
        "kpi": True,
        "drug_list": True,
        "drug_entry": True
    },

    "Pharmacist": {
        "pos": True,
        "stock": True,
        "dutylog": False,
        "admin_panel": False,
        "user_management": False,
        "inpatient": True,
        "kpi": True,
        "drug_list": False,
        "drug_entry": True
    },

    "Director": {
        "pos": True,
        "stock": True,
        "dutylog": True,
        "admin_panel": True,
        "user_management": True,
        "inpatient": True,
        "kpi": True,
        "drug_list": True,
        "drug_entry": True
    }
}

def has_permission(section):
    role = session.get("role")
    if not role:
        return False
    return ROLE_PERMISSIONS.get(role, {}).get(section, False)

@app.context_processor
def inject_permissions():
    return dict(has_permission=has_permission)
# ---------------------------------------------------------
# DB HELPERS
# ---------------------------------------------------------
def get_conn():
    return sqlite3.connect(DB_PATH)


def db_query(sql, params=()):
    conn = get_conn()
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return [dict(r) for r in rows]


def db_execute(sql, params=()):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(sql, params)
    conn.commit()
    conn.close()


def db_insert_many(sql, rows):
    conn = get_conn()
    cur = conn.cursor()
    cur.executemany(sql, rows)
    conn.commit()
    conn.close()
# ---------------------------------------------------------
# CREATE SQLITE TABLES
# ---------------------------------------------------------
conn = sqlite3.connect("pharmacy.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS LoginAttempts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT,
    attempts INTEGER,
    last_attempt TEXT,
    locked INTEGER DEFAULT 0
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS DrugList (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    drug_name TEXT,
    sales_price REAL,
    starting_balance REAL,
    categories TEXT,
    cost_price REAL
)
""")

cursor.execute("""
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
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS Transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    drug TEXT,
    qty REAL,
    price REAL,
    total REAL,
    cost REAL,
    cost_total REAL,
    customer_type TEXT,
    patient TEXT,
    cashier TEXT,
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
)
""")

cursor.execute("""
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
)
""")

cursor.execute("""
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
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS DutyLog (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT,
    login_time TEXT,
    logout_time TEXT
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS Users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    [User Id] TEXT UNIQUE,
    Password TEXT,
    Name TEXT,
    Role TEXT,
    Status TEXT DEFAULT 'Active'
)
""")

cursor.execute("""
CREATE TABLE IF NOT EXISTS AuditTrail (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    [User Id] TEXT,
    Action TEXT,
    Details TEXT,
    Timestamp TEXT
)
""")


conn.commit()
conn.close()
# ---------------------------------------------------------
# SETTINGS BACKEND
# ---------------------------------------------------------
def load_system_settings():
    conn = get_conn()
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT * FROM Settings LIMIT 1")
    row = cur.fetchone()

    if not row:
        cur.execute("""
        INSERT INTO Settings (hospital_name, auto_backup, backup_freq)
        VALUES (?, ?, ?)
        """, ("Ibukunolu Hospital", "on", "daily"))
        conn.commit()
        conn.close()
        return {
            "hospital_name": "Ibukunolu Hospital",
            "auto_backup": "on",
            "backup_freq": "daily",
        }

    settings = {
        "hospital_name": row["hospital_name"],
        "auto_backup": row["auto_backup"],
        "backup_freq": row["backup_freq"],
    }
    conn.close()
    return settings


def save_system_settings(data):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
    UPDATE Settings
    SET hospital_name = ?, auto_backup = ?, backup_freq = ?
    WHERE id = 1
    """, (data["hospital_name"], data["auto_backup"], data["backup_freq"]))
    conn.commit()
    conn.close()
# ---------------------------------------------------------
# SQLITE HELPERS
# ---------------------------------------------------------
def db_query(sql, params=()):
    conn = sqlite3.connect("pharmacy.db")
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()
    return [dict(row) for row in rows]

def db_execute(sql, params=()):
    conn = sqlite3.connect("pharmacy.db")
    cur = conn.cursor()
    cur.execute(sql, params)
    conn.commit()
    conn.close()

def db_insert_many(sql, rows):
    conn = sqlite3.connect("pharmacy.db")
    cur = conn.cursor()
    cur.executemany(sql, rows)
    conn.commit()
    conn.close()

# ---------------------------------------------------------
# EXCEL LOADER (ONLY FOR MIGRATION)
# ---------------------------------------------------------
def load_sheet(sheet_name):
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
        df.columns = df.columns.astype(str).str.strip()
        return df
    except:
        return pd.DataFrame()

# ---------------------------------------------------------
# INITIAL MIGRATION FROM EXCEL TO SQLITE USERS
# ---------------------------------------------------------
def migrate_users_from_excel():
    df = load_sheet("Users Data")
    if df.empty:
        return

    df.columns = df.columns.astype(str).str.strip()

    required_cols = {"User Id", "Password", "Name", "Role"}
    if not required_cols.issubset(set(df.columns)):
        return

    for _, row in df.iterrows():
        user_id = str(row["User Id"]).strip()
        name = str(row["Name"]).strip()
        role = str(row["Role"]).strip()
        raw_password = str(row["Password"]).strip()

        existing = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
        if existing:
            continue

        hashed = bcrypt.hashpw(raw_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

        db_execute("""
    INSERT INTO Users ([User Id], Password, Name, Role, Status)
    VALUES (?, ?, ?, ?, 'Active')
""", (user_id, hashed, name, role))

migrate_users_from_excel()



# ---------------------------------------------------------
# LOGIN ATTEMPT HELPERS
# ---------------------------------------------------------
def get_attempts(user_id):
    rows = db_query("SELECT * FROM LoginAttempts WHERE user_id = ?", (user_id,))
    return rows[0] if rows else None

def update_attempts(user_id, success=False):
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    record = get_attempts(user_id)

    if success:
        db_execute("DELETE FROM LoginAttempts WHERE user_id = ?", (user_id,))
        return

    if not record:
        db_execute("""
            INSERT INTO LoginAttempts (user_id, attempts, last_attempt, locked)
            VALUES (?, ?, ?, ?)
        """, (user_id, 1, now, 0))
        return

    attempts = record["attempts"] + 1
    locked = 1 if attempts >= 5 else 0

    db_execute("""
        UPDATE LoginAttempts
        SET attempts = ?, last_attempt = ?, locked = ?
        WHERE user_id = ?
    """, (attempts, now, locked, user_id))

# ---------------------------------------------------------
# LOGIN ROUTE (USING SQLITE + BCRYPT)
# ---------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user_id = request.form.get("username", "").strip()
        password_input = request.form.get("password", "").strip()
        role_input = request.form.get("role", "").strip()

        record = get_attempts(user_id)
        if record and record["locked"] == 1:
            return render_template("login.html",
                                   error="Account locked due to too many failed attempts.")

        rows = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
        if not rows:
            update_attempts(user_id, success=False)
            return render_template("login.html",
                                   error="Invalid User Id or Role.")

        user_row = rows[0]

        # 🔥 FIXED DEACTIVATION CHECK
        if user_row["Password"] == "DEACTIVATED":
            return render_template("login.html",
                                   error="Account is deactivated.")

        # Role check
        if user_row["Role"].strip() != role_input.strip():
            update_attempts(user_id, success=False)
            return render_template("login.html",
                                   error="Invalid User Id or Role.")

        stored_hash = user_row["Password"]

        # Password check
        if not bcrypt.checkpw(password_input.encode("utf-8"), stored_hash.encode("utf-8")):
            update_attempts(user_id, success=False)
            return render_template("login.html",
                                   error="Incorrect password.")

        update_attempts(user_id, success=True)

        # Session setup
        session["user"] = user_row["Name"]
        session["username"] = user_row["Name"]
        session["role"] = user_row["Role"]
        session["user_id"] = user_row["User Id"]
        session["login_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Duty log
        db_execute("""
            INSERT INTO DutyLog (user_id, login_time, logout_time)
            VALUES (?, ?, ?)
        """, (session["user_id"], session["login_time"], None))

        return redirect("/dashboard")

    return render_template("login.html")

# ---------------------------------------------------------
# DASHBOARD
# ---------------------------------------------------------
@app.route("/dashboard")
def dashboard():
    if "user" not in session:
        return redirect("/login")

    return render_template(
        "dashboard.html",
        user=session.get("user"),
        role=session.get("role")
    )

# ---------------------------------------------------------
# LOGOUT
# ---------------------------------------------------------
@app.route("/logout")
def logout():
    if "user_id" in session:
        logout_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        db_execute("""
            UPDATE DutyLog
            SET logout_time = ?
            WHERE user_id = ?
            AND logout_time IS NULL
        """, (logout_time, session["user_id"]))

    session.clear()
    return redirect("/login")

# ---------------------------------------------------------
# POS PAGE
# ---------------------------------------------------------
@app.route("/pos", methods=["GET"])
def pos():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("pos"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    drugs = db_query("SELECT * FROM DrugList")
    stock = db_query("SELECT * FROM Stock")

    return render_template(
        "pos.html",
        user=session.get("user"),
        role=session.get("role"),
        drugs=drugs,
        stock=stock,
        patient=""
    )

@app.route("/pos/success")
def pos_success():
    cart = session.get("last_cart", [])
    return render_template("pos_success.html", cart=cart)

def generate_transaction_id():
    today = datetime.now().strftime("%Y%m%d")
    today_sql = datetime.now().strftime("%Y-%m-%d")

    conn = sqlite3.connect("pharmacy.db")
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) AS count FROM Stock WHERE date = ?", (today_sql,))
    count = cur.fetchone()["count"]

    conn.close()

    suffix = count + 1
    return f"{today}-D{suffix}"

@app.route("/pos/submit", methods=["POST"])
def pos_submit():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("pos"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    role = session.get("role")
    if role in ["Director", "Administrator", "Auditor"]:
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    # Implement POS transaction logic here as needed
    return redirect("/pos/success")

# ---------------------------------------------------------
# STOCK VIEW
# ---------------------------------------------------------
@app.route("/stock")
def stock_view():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("stock"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    user = session["user"]
    role = session.get("role", "User")

    conn = sqlite3.connect("pharmacy.db")
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("""
        SELECT 
            date, time, drug_name, categories, opening_balance,
            transaction_type, quantity_sold, quantity_remaining,
            cost_price, sales_price, customer_type, transaction_id,
            remark, bill, sales_price_sum, sum_cost_value
        FROM Stock
        ORDER BY date DESC, time DESC
    """)

    stock_rows = cur.fetchall()
    conn.close()

    return render_template(
        "stockview.html",
        user=user,
        role=role,
        stock=stock_rows
    )

# ---------------------------------------------------------
# DRUG LIST
# ---------------------------------------------------------
@app.route("/druglist")
def druglist():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("drug_list"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    drugs = db_query("SELECT * FROM DrugList")

    return render_template(
        "druglist.html",
        user=session.get("user"),
        role=session.get("role"),
        drugs=drugs
    )

# ---------------------------------------------------------
# DRUG ENTRY
# ---------------------------------------------------------
@app.route("/drug_entry", methods=["GET", "POST"])
def drug_entry():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("drug_entry"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    if request.method == "POST":
        role = session.get("role")
        if role in ["Director", "Administrator"]:
            return render_template("no_access.html",
                user=session.get("user"),
                role=session.get("role")
            ), 403

        drug_name = request.form.get("drug_name")
        sales_price = request.form.get("sales_price")
        starting_balance = request.form.get("starting_balance")
        categories = request.form.get("categories")
        cost_price = request.form.get("cost_price")

        db_execute("""
            INSERT INTO DrugList (drug_name, sales_price, starting_balance, categories, cost_price)
            VALUES (?, ?, ?, ?, ?)
        """, (drug_name, sales_price, starting_balance, categories, cost_price))

        return redirect("/druglist")

    return render_template(
        "drug_entry.html",
        user=session.get("user"),
        role=session.get("role")
    )

# ---------------------------------------------------------
# INPATIENT DATA
# ---------------------------------------------------------
@app.route("/inpatient_data")
def inpatient_data():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("inpatient"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    rows = db_query("SELECT * FROM InpatientData ORDER BY date DESC")

    return render_template(
        "inpatient_data.html",
        user=session.get("user"),
        role=session.get("role"),
        data=rows
    )

# ---------------------------------------------------------
# KPI DASHBOARD
# ---------------------------------------------------------
@app.route("/kpi_dashboard")
def kpi_dashboard():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("kpi"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    return render_template(
        "kpi_dashboard.html",
        user=session.get("user"),
        role=session.get("role")
    )

# ---------------------------------------------------------
# INPATIENT BILLING
# ---------------------------------------------------------
@app.route("/inpatient/bill/calculator", methods=["GET"])
def inpatient_bill_page():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("inpatient"):
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    patients = db_query("SELECT DISTINCT patient_name FROM InpatientData ORDER BY patient_name ASC")
    patients = [p["patient_name"] for p in patients]

    selected_patient = request.args.get("patient", "")

    return render_template(
        "bill_calculator.html",
        user=session.get("user"),
        role=session.get("role"),
        patients=patients,
        selected_patient=selected_patient
    )

@app.route("/inpatient/bill/get_dates")
def inpatient_get_dates():
    patient = request.args.get("patient")

    if not patient:
        return {"dates": []}

    rows = db_query("""
        SELECT DISTINCT date
        FROM InpatientData
        WHERE LOWER(patient_name) = LOWER(?)
        ORDER BY date ASC
    """, (patient,))

    dates = [r["date"] for r in rows]

    return {"dates": dates}

@app.route("/inpatient/bill/calculate", methods=["POST"])
def inpatient_bill_calculate():
    if "user" not in session:
        return redirect("/login")

    patient = request.form.get("patient")
    admission = request.form.get("admission_date")
    discharge = request.form.get("discharge_date")

    rows = db_query("""
        SELECT *
        FROM InpatientData
        WHERE LOWER(patient_name) = LOWER(?)
        AND date >= ?
        AND date <= ?
        ORDER BY date ASC
    """, (patient, admission, discharge))

    total_bill = sum([r["bill_total"] for r in rows])
    total_qty = sum([r["qty_sold"] for r in rows])

    import datetime as dt
    generated_date = dt.datetime.now().strftime("%d-%b-%Y %I:%M %p")
    invoice_number = f"INV-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}"

    return render_template(
        "bill_result.html",
        user=session.get("user"),
        role=session.get("role"),
        patient=patient,
        admission_date=admission,
        discharge_date=discharge,
        bill_items=rows,
        total_bill=total_bill,
        total_qty=total_qty,
        generated_date=generated_date,
        invoice_number=invoice_number
    )

@app.route("/inpatient/history")
def inpatient_history():
    patient = request.args.get("patient")

    rows = db_query("""
        SELECT *
        FROM InpatientData
        WHERE patient_name = ?
        ORDER BY date ASC
    """, (patient,))

    return render_template(
        "inpatient_history.html",
        user=session.get("user"),
        role=session.get("role"),
        patient=patient,
        rows=rows
    )

@app.route("/inpatient/bill/receipt")
def inpatient_bill_receipt():
    patient = request.args.get("patient")
    admission = request.args.get("admission")
    discharge = request.args.get("discharge")

    rows = db_query("""
        SELECT *
        FROM InpatientData
        WHERE LOWER(patient_name) = LOWER(?)
        AND date >= ?
        AND date <= ?
        ORDER BY date ASC
    """, (patient, admission, discharge))

    total_bill = sum([r["bill_total"] for r in rows])

    import datetime as dt
    generated_date = dt.datetime.now().strftime("%d-%b-%Y %I:%M %p")
    invoice_number = f"INV-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}"

    return render_template(
        "bill_receipt.html",
        patient=patient,
        admission_date=admission,
        discharge_date=discharge,
        bill_items=rows,
        total_bill=total_bill,
        generated_date=generated_date,
        invoice_number=invoice_number,
        user=session.get("user")
    )

# ---------------------------------------------------------
# DRUG RESTOCK
# ---------------------------------------------------------
@app.route("/drug/restock/save", methods=["POST"])
def drug_restock_save():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if role in ["Director", "Administrator"]:
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    data = request.get_json()
    items = data["items"]

    import datetime as dt

    for item in items:
        transaction_id = "STK" + dt.datetime.now().strftime("%Y%m%d%H%M%S")

        existing = db_query("""
            SELECT quantity_remaining
            FROM Stock
            WHERE LOWER(drug_name) = LOWER(?)
            ORDER BY id DESC
            LIMIT 1
        """, (item["drug_name"],))

        opening_balance = existing[0]["quantity_remaining"] if existing else 0

        qty_added = int(item["qty_added"])
        quantity_remaining = opening_balance + qty_added

        db_execute("""
            INSERT INTO Stock (
                date, time, drug_name, categories, opening_balance,
                transaction_type, quantity_sold, quantity_remaining,
                cost_price, sales_price, customer_type, transaction_id,
                remark, bill, sales_price_sum, sum_cost_value
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            item["date"],
            dt.datetime.now().strftime("%H:%M:%S"),
            item["drug_name"],
            item["category"],
            opening_balance,
            "Restocked",
            0,
            quantity_remaining,
            item["cost_price"],
            item["sales_price"],
            "Stock",
            transaction_id,
            "Restocked",
            0,
            0,
            0
        ))

    return {"status": "success"}

# ---------------------------------------------------------
# DRUGLIST APIs
# ---------------------------------------------------------
@app.route("/api/druglist")
def api_druglist():
    rows = db_query("""
        SELECT 
            drug_name,
            categories AS category,
            sales_price,
            cost_price,
            starting_balance
        FROM DrugList
        ORDER BY drug_name ASC
    """)
    return {"drugs": rows}

@app.route("/druglist/update", methods=["POST"])
def druglist_update():
    data = request.get_json()
    db_execute("""
        UPDATE DrugList
        SET categories = ?, sales_price = ?, cost_price = ?, starting_balance = ?
        WHERE id = ?
    """, (data["category"], data["sales_price"], data["cost_price"],
          data["starting_balance"], data["id"]))
    return {"status": "success"}

@app.route("/druglist/create", methods=["POST"])
def druglist_create():
    data = request.get_json()
    db_execute("""
        INSERT INTO DrugList (drug_name, categories, sales_price, cost_price, starting_balance)
        VALUES (?, ?, ?, ?, ?)
    """, (data["drug_name"], data["category"], data["sales_price"],
          data["cost_price"], data["starting_balance"]))
    return {"status": "success"}

@app.route("/druglist/delete", methods=["POST"])
def druglist_delete():
    data = request.get_json()
    db_execute("DELETE FROM DrugList WHERE id = ?", (data["id"],))
    return {"status": "success"}

# ---------------------------------------------------------
# PROFILE (NO PICTURE, HOSPITAL LOGO USED IN HTML)
# ---------------------------------------------------------
@app.route("/profile", methods=["GET", "POST"])
def profile():
    if "user" not in session:
        return redirect("/login")

    user_id = session.get("user_id")
    role = session.get("role")

    user_rows = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
    if not user_rows:
        return render_template("profile.html",
                               error="User record not found.",
                               user_data={})

    user_data = user_rows[0]

    error = None
    success = None

    if request.method == "POST":
        new_name = request.form.get("name")
        new_password = request.form.get("password")

        if new_password:
            hashed = bcrypt.hashpw(new_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
            db_execute("""
                UPDATE Users
                SET Password = ?
                WHERE [User Id] = ?
            """, (hashed, user_id))

        db_execute("""
            UPDATE Users
            SET Name = ?
            WHERE [User Id] = ?
        """, (new_name, user_id))

        session["user"] = new_name
        success = "Profile updated successfully."

        user_rows = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
        user_data = user_rows[0]

    return render_template(
        "profile.html",
        user_data=user_data,
        role=role,
        error=error,
        success=success
    )

# ---------------------------------------------------------
# ADMIN PANEL
# ---------------------------------------------------------
@app.route("/admin/dashboard")
def admin_dashboard():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("admin_panel"):
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    return render_template(
        "admin_dashboard.html",
        user=session.get("user"),
        role=session.get("role")
    )

@app.route("/admin/users")
def admin_users():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")

    if not has_permission("admin_panel") or not has_permission("user_management"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=role
        ), 403

    users = db_query("SELECT [User Id], Password, Name, Role FROM Users ORDER BY Name ASC")

    return render_template(
        "admin_users.html",
        user=session.get("user"),
        role=role,
        users=users
    )

@app.route("/admin/users/add", methods=["POST"])
def admin_users_add():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if not has_permission("admin_panel") or not has_permission("user_management"):
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=role
        ), 403

    # Get form fields
    name = request.form.get("name")
    user_id = request.form.get("User ID")   # ✔ matches your HTML
    password = request.form.get("password")
    user_role = request.form.get("role")

    # Validate empty User ID
    if not user_id or user_id.strip() == "":
        users = db_query("SELECT [User Id], Password, Name, Role FROM Users ORDER BY Name ASC")
        return render_template(
            "admin_users.html",
            user=session.get("user"),
            role=role,
            users=users,
            error="User ID cannot be empty."
        )

    # ✔ CHECK IF USER ALREADY EXISTS
    existing = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
    if existing:
        users = db_query("SELECT [User Id], Password, Name, Role FROM Users ORDER BY Name ASC")
        return render_template(
            "admin_users.html",
            user=session.get("user"),
            role=role,
            users=users,
            error="User already exists."
        )

    # Hash password
    hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

    # Insert new user
    db_execute("""
        INSERT INTO Users ([User Id], Password, Name, Role)
        VALUES (?, ?, ?, ?)
    """, (user_id, hashed, name, user_role))

    # Reload user list
    users = db_query("SELECT [User Id], Password, Name, Role FROM Users ORDER BY Name ASC")

    return render_template(
        "admin_users.html",
        user=session.get("user"),
        role=role,
        users=users,
        message="User added successfully."
    )



@app.route("/admin/users/edit", methods=["POST"])
def admin_users_edit():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if not has_permission("admin_panel") or not has_permission("user_management"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=role
        ), 403

    user_id = request.form.get("userid")
    name = request.form.get("name")
    password = request.form.get("password")
    user_role = request.form.get("role")

    if password:
        hashed = bcrypt.hashpw(password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        db_execute("""
            UPDATE Users
            SET Password = ?
            WHERE [User Id] = ?
        """, (hashed, user_id))

    db_execute("""
        UPDATE Users
        SET Name = ?, Role = ?
        WHERE [User Id] = ?
    """, (name, user_role, user_id))

    users = db_query("SELECT [User Id], Password, Name, Role FROM Users ORDER BY Name ASC")

    return render_template(
        "admin_users.html",
        user=session.get("user"),
        role=role,
        users=users,
        message="User updated successfully."
    )

@app.route("/admin/users/delete")
def admin_users_delete():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if role != "Admin":
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=role
        ), 403

    user_id = request.args.get("user")
    if not user_id:
        return redirect("/admin/users?error=Invalid+User")

    # DELETE USER
    db_execute("DELETE FROM Users WHERE [User Id] = ?", (user_id,))

    # ⭐ No audit trail here — fully removed

    return redirect("/admin/users?message=User+deleted+successfully")


# ---------------------------------------------------------
# DEACTIVATE USER
# ---------------------------------------------------------
@app.route("/admin/users/deactivate")
def admin_users_deactivate():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if role != "Admin":
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=role
        ), 403

    user_id = request.args.get("user")
    if not user_id:
        return redirect("/admin/users?message=Invalid+User")

    # Disable login by setting password to a non-usable value
    db_execute("UPDATE Users SET Password='DEACTIVATED' WHERE [User Id] = ?", (user_id,))

    users = db_query("SELECT [User Id], Password, Name, Role FROM Users ORDER BY Name ASC")

    return render_template(
        "admin_users.html",
        user=session.get("user"),
        role=role,
        users=users,
        message="User deactivated successfully."
    )

@app.route("/admin/users/reactivate")
def admin_users_reactivate():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if role != "Admin":
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=role
        ), 403

    user_id = request.args.get("user")
    if not user_id:
        return redirect("/admin/users?error=Invalid+User")

    user = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
    if not user:
        return redirect("/admin/users?error=User+not+found")

    user = user[0]

    # ✔ Already active?
    if user["Password"] != "DEACTIVATED":
        return redirect("/admin/users?error=User+is+already+active")

    # ✔ Reactivate user (set new default password)
    new_password = bcrypt.hashpw("12345".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

    db_execute("""
        UPDATE Users SET Password=? WHERE [User Id]=?
    """, (new_password, user_id))

    # AUDIT TRAIL
    db_execute("""
        INSERT INTO AuditTrail (action, user_id, timestamp)
        VALUES (?, ?, ?)
    """, (f"Reactivated user {user_id}", session.get("user_id"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")))

    return redirect("/admin/users?message=User+reactivated+successfully")

@app.route("/admin/logs")
def admin_logs():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if not has_permission("admin_panel") or not has_permission("dutylog"):
        return render_template("no_access.html",
                               user=session.get("user"),
                               role=role), 403

    logs = db_query("SELECT * FROM DutyLog ORDER BY login_time DESC")
    return render_template(
        "admin_logs.html",
        user=session.get("user"),
        role=role,
        logs=logs,
    )


@app.route("/admin/sheets")
def admin_sheets():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if not has_permission("admin_panel"):
        return render_template("no_access.html",
                               user=session.get("user"),
                               role=role), 403

    return render_template(
        "admin_sheets.html",
        user=session.get("user"),
        role=role,
    )


from kpi_manager import load_kpi_settings

@app.route("/admin/kpi")
def admin_kpi():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if not has_permission("admin_panel"):
        return render_template("no_access.html",
                               user=session.get("user"),
                               role=role), 403

    settings = load_kpi_settings()

    return render_template(
        "admin_kpi.html",
        user=session.get("user"),
        role=role,
        settings=settings,
    )


from kpi_manager import save_kpi_settings

@app.route("/admin/kpi/save", methods=["POST"])
def admin_kpi_save():
    if "user" not in session:
        return jsonify({"status": "error", "message": "Not logged in"})

    role = session.get("role")
    if role not in ["Director", "Administrator", "Admin"]:
        return jsonify({"status": "error", "message": "Access denied"})

    data = request.get_json()

    save_kpi_settings(data)

    return jsonify({"status": "success", "message": "KPI settings saved"})


@app.route("/admin/settings")
def admin_settings():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if role not in ["Director", "Administrator", "Admin"]:
        return render_template("no_access.html",
                               user=session.get("user"),
                               role=role), 403

    settings = load_system_settings()
    return render_template(
        "admin_settings.html",
        user=session.get("user"),
        role=role,
        settings=settings,
    )


@app.route("/admin/settings/save", methods=["POST"])
def admin_settings_save():
    if "user" not in session:
        return jsonify({"status": "error", "message": "Not logged in"})

    role = session.get("role")
    if role not in ["Director", "Administrator", "Admin"]:
        return jsonify({"status": "error", "message": "Access denied"})

    data = request.get_json()
    save_system_settings(data)
    return jsonify({"status": "success", "message": "Settings saved"})


@app.route("/admin/backups")
def admin_backups():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if not has_permission("admin_panel"):
        return render_template("no_access.html",
                               user=session.get("user"),
                               role=role), 403

    # You can later add real backup listing here
    backups = []
    return render_template(
        "admin_backups.html",
        user=session.get("user"),
        role=role,
        backups=backups,
    )
@app.route("/admin/users/activity/export/<user_id>")
def export_user_activity(user_id):
    logs = db_query("""
        SELECT login_time, logout_time
        FROM DutyLog
        WHERE user_id = ?
        ORDER BY login_time DESC
    """, (user_id,))

    import pandas as pd
    df = pd.DataFrame(logs)

    filename = f"user_activity_{user_id}.xlsx"
    filepath = os.path.join("exports", filename)

    os.makedirs("exports", exist_ok=True)
    df.to_excel(filepath, index=False)

    return send_file(filepath, as_attachment=True)

@app.route("/admin/backups/restore")
def admin_restore_backup():
    file = request.args.get("file")

    enc_path = os.path.join(BACKUP_DIR, file)

    try:
        restore_backup(enc_path)
        return redirect("/admin/backups?message=Backup+restored+successfully")
    except Exception as e:
        print("Restore error:", e)
        return redirect("/admin/backups?error=Restore+failed")
@app.route("/admin/backups/cloud-status")
def cloud_status():
    local_stats = get_local_backup_stats()
    cloud_backups = list_cloud_backups()

    cloud_ok = cloud_backups is not None
    cloud_count = len(cloud_backups) if cloud_ok else 0

    return render_template(
        "admin_cloud_status.html",
        local_stats=local_stats,
        cloud_ok=cloud_ok,
        cloud_count=cloud_count,
        cloud_backups=cloud_backups
    )

@app.route("/admin/backups/cloud-download")
def cloud_download():
    file_id = request.args.get("id")
    name = request.args.get("name")

    save_path = os.path.join(BACKUP_DIR, name)

    try:
        download_backup_from_drive(file_id, save_path)
        return send_file(save_path, as_attachment=True)
    except Exception as e:
        print("Cloud download error:", e)
        return redirect("/admin/backups/cloud-status?error=Download+failed")

@app.route("/admin/backups/cloud-restore")
def cloud_restore():
    file_id = request.args.get("id")
    name = request.args.get("name")

    save_path = os.path.join(BACKUP_DIR, name)

    try:
        download_backup_from_drive(file_id, save_path)
        restore_backup(save_path)
        return redirect("/admin/backups/cloud-status?message=Cloud+restore+successful")
    except Exception as e:
        print("Cloud restore error:", e)
        return redirect("/admin/backups/cloud-status?error=Cloud+restore+failed")


@app.route("/admin/backups/create")
def admin_backups_create():
    try:
        create_backup()
        return redirect("/admin/backups?message=Backup+created+successfully")
    except Exception as e:
        print("Backup error:", e)
        return redirect("/admin/backups?error=Failed+to+create+backup")

@app.route("/admin/backups/download")
def download_backup():
    file = request.args.get("file")
    return send_file(os.path.join(BACKUP_DIR, file), as_attachment=True)

def auto_backup_scheduler():
    while True:
        run_auto_backup_if_due()
        time.sleep(60)

Thread(target=auto_backup_scheduler, daemon=True).start()
@app.route("/admin/backups/history")
def backup_history():
    local = get_local_backups()
    cloud = get_cloud_backups()

    all_backups = local + cloud

    # Sort newest first
    all_backups = sorted(all_backups, key=lambda x: x["date"], reverse=True)

    return render_template("admin_backup_history.html", backups=all_backups)
@app.route("/admin/backups/download")
def download_local_backup():
    file = request.args.get("file")
    path = os.path.join(BACKUP_DIR, file)

    return send_file(path, as_attachment=True)
@app.route("/admin/backups/restore")
def restore_local_backup():
    file = request.args.get("file")
    path = os.path.join(BACKUP_DIR, file)

    try:
        restore_backup(path)
        return redirect("/admin/backups/history?message=Restore+successful")
    except Exception:
        return redirect("/admin/backups/history?error=Restore+failed")
@app.route("/admin/backups/restore-wizard")
def restore_wizard():
    local = get_local_backups()
    cloud = get_cloud_backups()

    backups = local + cloud
    backups = sorted(backups, key=lambda x: x["date"], reverse=True)

    return render_template("admin_restore_wizard.html", backups=backups)

@app.route("/admin/backups/restore-run")
def restore_run():
    source = request.args.get("source")
    name = request.args.get("name")
    file_id = request.args.get("id")

    enc_path = os.path.join(BACKUP_DIR, name)

    try:
        # If cloud backup → download first
        if source == "cloud":
            download_backup_from_drive(file_id, enc_path)

        # Restore
        restore_backup(enc_path)

        return render_template("admin_restore_success.html", backup=name)

    except Exception as e:
        print("Restore error:", e)
        return render_template("admin_restore_failed.html", backup=name)

@app.route("/admin/backups/integrity")
def backup_integrity():
    results = scan_backup_integrity()
    return render_template("admin_integrity_scanner.html", results=results)

@app.route("/admin/backups/scheduler")
def backup_scheduler():
    status = get_scheduler_status()
    return render_template("admin_scheduler_dashboard.html", status=status)

@app.route("/admin/backups/cloud-monitor")
def cloud_monitor():
    status = get_cloud_sync_status()
    return render_template("admin_cloud_monitor.html", status=status)

import sqlite3, os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pharmacy.db")

def get_kpis():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # Example queries — adjust to your actual DB schema
    cur.execute("SELECT SUM(total_amount) FROM sales WHERE date = DATE('now')")
    sales_today = cur.fetchone()[0] or 0

    cur.execute("SELECT SUM(total_amount) FROM sales WHERE strftime('%Y-%m', date) = strftime('%Y-%m', 'now')")
    sales_month = cur.fetchone()[0] or 0

    cur.execute("SELECT COUNT(*) FROM sales")
    total_transactions = cur.fetchone()[0]

    cur.execute("SELECT COUNT(*) FROM customers")
    total_customers = cur.fetchone()[0]

    cur.execute("""
        SELECT product_name, SUM(quantity) AS qty 
        FROM sales_items 
        GROUP BY product_name 
        ORDER BY qty DESC 
        LIMIT 5
    """)
    top_products = cur.fetchall()

    cur.execute("SELECT product_name, stock FROM products WHERE stock < 10")
    low_stock = cur.fetchall()

    conn.close()

    return {
        "sales_today": sales_today,
        "sales_month": sales_month,
        "transactions": total_transactions,
        "customers": total_customers,
        "top_products": top_products,
        "low_stock": low_stock
    }
from log_manager import get_logs

@app.route("/admin/system-logs")
def admin_system_logs():
    logs = get_logs()
    logs = sorted(logs, key=lambda x: x["time"], reverse=True)
    return render_template("admin_logs.html", logs=logs)

from flask import redirect

@app.route("/")
def home():
    return redirect("/login")

# ---------------------------------------------------------
# RUN APP
# ---------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
