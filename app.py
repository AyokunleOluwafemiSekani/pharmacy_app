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
import user_agents
from kpi_manager import get_kpis

app = Flask(__name__)
app.secret_key = "super_secret_key"

# Excel file (ONLY used for initial migration)
EXCEL_PATH = os.path.join("static", "Ibukunolu new pharmacy database.xlsx")
DB_PATH = "pharmacynew.db"

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
        "low_stock": True,
        "inpatient_bill": True,
        "drug_list": True,
        "drug_entry": True
    },

    "Administrator": {
        "pos": False,
        "stock": True,
        "dutylog": True,
        "admin_panel": True,
        "user_management": True,
        "inpatient": True,
        "kpi": True,
        "low_stock": False,
        "inpatient_bill": False,
        "drug_list": False,
        "drug_entry": False
    },

    "HIM Officer": {
        "pos": False,
        "stock": False,
        "dutylog": False,
        "admin_panel": False,
        "user_management": False,
        "inpatient": True,
        "kpi": False,
        "low_stock": False,
        "inpatient_bill": True,
        "drug_list": False,
        "drug_entry": False
    },

    "Auditor": {
        "pos": True,
        "stock": True,
        "dutylog": False,
        "admin_panel": False,
        "user_management": False,
        "inpatient": True,
        "kpi": True,
        "low_stock": True,
        "inpatient_bill": False,
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
        "low_stock": True,
        "inpatient_bill": True,
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
        "low_stock": False,
        "inpatient_bill": False,
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

def db_execute(query, params=()):
    conn = sqlite3.connect(DB_PATH, timeout=5)
    cur = conn.cursor()
    cur.execute(query, params)
    conn.commit()
    conn.close()

def db_query(query, params=()):
    conn = sqlite3.connect(DB_PATH, timeout=5)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute(query, params)
    rows = cur.fetchall()
    conn.close()
    # ⭐ FIX: convert sqlite3.Row → dict
    return [dict(row) for row in rows]

def db_insert_many(sql, rows):
    conn = get_conn()
    cur = conn.cursor()
    cur.executemany(sql, rows)
    conn.commit()
    conn.close()

def enable_wal_mode():
    conn = sqlite3.connect(DB_PATH, timeout=5)
    cur = conn.cursor()
    cur.execute("PRAGMA journal_mode=WAL;")
    conn.commit()
    conn.close()

# ---------------------------------------------------------
# EXCEL LOADER
# ---------------------------------------------------------
def load_sheet(sheet_name):
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=sheet_name)
        df.columns = df.columns.astype(str).str.strip()
        return df
    except:
        return pd.DataFrame()

# ---------------------------------------------------------
# MIGRATIONS
# ---------------------------------------------------------
def migrate_users_from_excel():
    df = load_sheet("Users Data")
    if df.empty:
        print("Users Data sheet empty or missing.")
        return

    required_cols = {"User Id", "Password", "Name", "Role"}
    if not required_cols.issubset(df.columns):
        print("Users Data sheet missing required columns.")
        return

    rows = []
    for _, row in df.iterrows():
        user_id = str(row["User Id"]).strip()
        name = str(row["Name"]).strip()
        role = str(row["Role"]).strip()
        raw_password = str(row["Password"]).strip()

        existing = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
        if existing:
            continue

        hashed = bcrypt.hashpw(raw_password.encode("utf-8"), bcrypt.gensalt()).decode("utf-8")
        rows.append((user_id, hashed, name, role))

    db_insert_many("""
        INSERT INTO Users ([User Id], Password, Name, Role)
        VALUES (?, ?, ?, ?)
    """, rows)

    print("Users imported.")

def migrate_druglist_from_excel():
    df = load_sheet("Drug List")
    if df.empty:
        print("Drug List sheet empty or missing.")
        return

    rows = []
    for _, row in df.iterrows():
        rows.append((
            str(row["Drug Name"]).strip(),
            float(row["Sales Price"]),
            float(row["Starting Balance"]),
            str(row["Categories"]).strip(),
            float(row["Cost Price"])
        ))

    db_insert_many("""
        INSERT INTO DrugList (drug_name, sales_price, starting_balance, categories, cost_price)
        VALUES (?, ?, ?, ?, ?)
    """, rows)

    print("DrugList imported.")

def safe_float(value):
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except:
        return 0.0

def migrate_stock_from_excel():
    df = load_sheet("StockView")
    if df.empty:
        print("StockView sheet empty or missing.")
        return

    rows = []
    for _, row in df.iterrows():
        rows.append((
            str(row["DATE"]).strip(),
            str(row["TIME"]).strip(),
            str(row["DRUG NAME"]).strip(),
            str(row["CATEGORIES"]).strip(),
            safe_float(row["OPENING BALANCE"]),
            str(row["TRANSACTION TYPE"]).strip(),
            safe_float(row["QUANTITY SOLD"]),
            safe_float(row["QUANTITY REMAINING"]),
            safe_float(row["COST PRICE"]),
            safe_float(row["SALES PRICE"]),
            str(row["CUSTOMER TYPE"]).strip(),
            str(row["TRANSACTION ID"]).strip(),
            str(row["REMARK"]).strip(),
            safe_float(row["BILL"]),
            safe_float(row["Sum of Sales Value"]),
            safe_float(row["Sum of Cost Value"])
        ))

    db_insert_many("""
        INSERT INTO Stock (
            date, time, drug_name, categories, opening_balance,
            transaction_type, quantity_sold, quantity_remaining,
            cost_price, sales_price, customer_type, transaction_id,
            remark, bill, sales_price_sum, sum_cost_value
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)

    print("Stock imported.")

def generate_invoice_number():
    now = datetime.now()
    return "INV-" + now.strftime("%Y%m%d%H%M%S")


def migrate_inpatient_from_excel():
    df = load_sheet("Inpatient Data")

    if df.empty:
        print("InpatientData sheet empty or missing.")
        return

    rows = []

    for _, row in df.iterrows():
        rows.append((
            str(row["Date"]).strip(),
            str(row["Drug Name"]).strip(),
            str(row["Patient Name"]).strip(),
            str(row["Categories"]).strip(),
            float(row["Sales Price"]),
            str(row["Transaction ID"]).strip(),
            float(row["Bill Total"]),
            float(row["Qty Sold"]),
            generate_invoice_number()
        ))

    db_insert_many("""
        INSERT INTO InpatientData (
            date,
            drug_name,
            patient_name,
            categories,
            sales_price,
            transaction_id,
            bill_total,
            qty_sold,
            invoice_number
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)

    print("InpatientData imported")

def migrate_bill_from_excel():
    df = load_sheet("Bill")
    if df.empty:
        print("Bill sheet empty or missing.")
        return

    rows = []
    for _, row in df.iterrows():
        rows.append((
            str(row["Date"]).strip(),
            str(row["Drug Name"]).strip(),
            str(row["Patient Name"]).strip(),
            str(row["Categories"]).strip(),
            float(row["Sales Price"]),
            str(row["Transaction ID"]).strip(),
            float(row["Bill Total"]),
            str(row["Source"]).strip()
        ))

    db_insert_many("""
        INSERT INTO Bill (
            date, drug_name, patient_name, categories,
            sales_price, transaction_id, bill_total, source
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, rows)

    print("Bill imported.")

def migrate_dutylog_from_excel():
    df = load_sheet("DutyLog")
    if df.empty:
        print("DutyLog sheet empty or missing.")
        return

    rows = []
    for _, row in df.iterrows():
        logout = None if pd.isna(row["Logout Time"]) else str(row["Logout Time"]).strip()
        rows.append((
            str(row["User Id"]).strip(),
            str(row["Login Time"]).strip(),
            logout
        ))

    db_insert_many("""
        INSERT INTO DutyLog (user_id, login_time, logout_time)
        VALUES (?, ?, ?)
    """, rows)

    print("DutyLog imported.")

# ---------------------------------------------------------
# SAFE INITIALIZER (MUST BE AFTER MIGRATIONS)
# ---------------------------------------------------------
def initialize_database():
    fresh_db = not os.path.exists(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # ALWAYS ensure required tables exist
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
        qty_sold REAL,
        invoice_number TEXT
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
    CREATE TABLE IF NOT EXISTS activity_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        time TEXT,
        type TEXT,
        message TEXT,
        ip TEXT,
        user TEXT,
        role TEXT,
        device TEXT,
        browser TEXT,
        location TEXT
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

    # Run migrations ONLY on first creation
    if fresh_db:
        print("Fresh database created. Running Excel migrations...")
        migrate_users_from_excel()
        migrate_druglist_from_excel()
        migrate_stock_from_excel()
        migrate_inpatient_from_excel()
        migrate_bill_from_excel()
        migrate_dutylog_from_excel()
        print("Excel migration completed successfully.")

    enable_wal_mode()


# ---------------------------------------------------------
# RUN INITIALIZER (LAST)
# ---------------------------------------------------------
initialize_database()

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
# LOGIN ROUTE (USING SQLITE + BCRYPT + FULL LOGGING)
# ---------------------------------------------------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user_id = request.form.get("username", "").strip()
        password_input = request.form.get("password", "").strip()
        role_input = request.form.get("role", "").strip()

        ip = request.remote_addr
        user_agent_string = request.headers.get("User-Agent")

        # Check lockout
        record = get_attempts(user_id)
        if record and record["locked"] == 1:
            log_event("error",
                      f"Locked account attempted login: {user_id}",
                      ip, user_id, role_input, user_agent_string)
            return render_template("login.html",
                                   error="Account locked due to too many failed attempts.")

        # Fetch user
        rows = db_query("SELECT * FROM Users WHERE [User Id] = ?", (user_id,))
        if not rows:
            update_attempts(user_id, success=False)
            log_event("error",
                      f"Invalid User Id: {user_id}",
                      ip, user_id, role_input, user_agent_string)
            return render_template("login.html",
                                   error="Invalid User Id or Role.")

        user_row = rows[0]

        # Deactivated account
        if user_row["Password"] == "DEACTIVATED":
            log_event("error",
                      f"Deactivated account login attempt: {user_id}",
                      ip, user_id, role_input, user_agent_string)
            return render_template("login.html",
                                   error="Account is deactivated.")

        # Role mismatch
        if user_row["Role"].strip() != role_input.strip():
            update_attempts(user_id, success=False)
            log_event("error",
                      f"Role mismatch for {user_id}. Tried role: {role_input}",
                      ip, user_id, role_input, user_agent_string)
            return render_template("login.html",
                                   error="Invalid User Id or Role.")

        stored_hash = user_row["Password"]

        # Password mismatch
        if not bcrypt.checkpw(password_input.encode("utf-8"), stored_hash.encode("utf-8")):
            update_attempts(user_id, success=False)
            log_event("error",
                      f"Incorrect password for {user_id}",
                      ip, user_id, role_input, user_agent_string)
            return render_template("login.html",
                                   error="Incorrect password.")

        # Successful login
        update_attempts(user_id, success=True)

        session["user"] = user_row["Name"]
        session["username"] = user_row["Name"]
        session["role"] = user_row["Role"]
        session["user_id"] = user_row["User Id"]
        session["login_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Log successful login
        log_event("login",
                  f"{user_row['Name']} logged in successfully",
                  ip, user_row["Name"], user_row["Role"], user_agent_string)

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
# LOGOUT (WITH FULL LOGGING + DUTYLOG UPDATE)
# ---------------------------------------------------------
@app.route("/logout")
def logout():
    # Capture metadata
    ip = request.remote_addr
    ua = request.headers.get("User-Agent")

    user = session.get("user")
    role = session.get("role")
    user_id = session.get("user_id")

    # Log the logout event
    if user:
        log_event(
            "logout",
            f"{user} logged out",
            ip,
            user,
            role,
            ua
        )

    # Update DutyLog logout_time
    if user_id:
        logout_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        db_execute("""
            UPDATE DutyLog
            SET logout_time = ?
            WHERE user_id = ?
            AND logout_time IS NULL
        """, (logout_time, user_id))

    # Clear session
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

    drugs = db_query("SELECT * FROM DrugList ORDER BY drug_name ASC")
    stock = db_query("SELECT * FROM Stock ORDER BY id DESC")

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
    transaction_id = session.get("last_transaction_id", "N/A")
    return render_template("pos_success.html", cart=cart, transaction_id=transaction_id)


def generate_transaction_id():
    today_sql = datetime.now().strftime("%Y-%m-%d")
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) AS count FROM Stock WHERE date = ?", (today_sql,))
    count = cur.fetchone()["count"]
    conn.close()
    suffix = count + 1
    return f"{today_sql.replace('-', '')}-D{suffix}"


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
    if role in ["Director", "Administrator", "Auditor", "HIM Officer"]:
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    customer_type = request.form.get("customer_type", "OutPatient")
    patient = request.form.get("patient", "")
    cashier = session.get("user")

    # ⭐ Generate ONE transaction ID for the entire sale
    transaction_id = generate_transaction_id()
    session["last_transaction_id"] = transaction_id

    cart = []  # store items for success page

    # Loop through POS rows (1–100)
    for i in range(1, 101):
        drug = request.form.get(f"drug_{i}")
        qty = request.form.get(f"qty_{i}")
        price = request.form.get(f"price_{i}")
        total = request.form.get(f"total_{i}")

        if not drug or not qty or float(qty) <= 0:
            continue

        qty = float(qty)
        price = float(price)
        total = float(total)

        cart.append({
            "drug": drug,
            "qty": qty,
            "total": total
        })

        # Fetch current stock
        stock_row = db_query("SELECT * FROM DrugList WHERE drug_name = ?", (drug,))
        if not stock_row:
            continue

        stock_row = stock_row[0]
        opening_balance = float(stock_row["starting_balance"])
        remaining = opening_balance - qty

        # Update DrugList balance
        db_execute("""
            UPDATE DrugList SET starting_balance = ?
            WHERE drug_name = ?
        """, (remaining, drug))

        # Insert into Stock
        db_execute("""
            INSERT INTO Stock (
                date, time, drug_name, categories, opening_balance,
                transaction_type, quantity_sold, quantity_remaining,
                cost_price, sales_price, customer_type, transaction_id,
                remark, bill, sales_price_sum, sum_cost_value
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            datetime.now().strftime("%Y-%m-%d"),
            datetime.now().strftime("%H:%M:%S"),
            drug,
            stock_row["categories"],
            opening_balance,
            "Dispensed",
            qty,
            remaining,
            stock_row["cost_price"],
            price,
            customer_type,
            transaction_id,
            "",
            total,
            total,
            stock_row["cost_price"] * qty
        ))

        # Insert into Transactions
        db_execute("""
            INSERT INTO Transactions (
                drug, qty, price, total, cost, cost_total,
                customer_type, patient, cashier
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            drug, qty, price, total,
            stock_row["cost_price"],
            stock_row["cost_price"] * qty,
            customer_type, patient, cashier
        ))

        # ⭐ FLEXIBLE INPATIENT LOGIC
        if "inpatient" in customer_type.lower():
            db_execute("""
                INSERT INTO InpatientData (
                    date, drug_name, patient_name, categories,
                    sales_price, transaction_id, bill_total,
                    qty_sold, invoice_number
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                datetime.now().strftime("%Y-%m-%d"),
                drug,
                patient,
                stock_row["categories"],
                price,
                transaction_id,
                total,
                qty,
                transaction_id
            ))

    session["last_cart"] = cart
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

    conn = sqlite3.connect(DB_PATH)
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
# LOW STOCK PAGE WITH DATE FILTER + PDF EXPORT
# ---------------------------------------------------------
import datetime as dt
import io
import base64

@app.route("/stock/low", methods=["GET", "POST"])
def stock_low():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("low_stock"):
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    # Critical low stock threshold
    threshold = 20

    # Filters
    search = ""
    sort = "quantity_remaining"
    start_date = None
    end_date = None

    # -----------------------------
    # POST → Date filter
    # -----------------------------
    if request.method == "POST":
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")

        query = """
            SELECT *
            FROM Stock
            WHERE quantity_remaining < ?
            AND date >= ?
            AND date <= ?
        """
        params = [threshold, start_date, end_date]

    # -----------------------------
    # GET → Search + Sort
    # -----------------------------
    else:
        search = request.args.get("search", "").strip().lower()
        sort = request.args.get("sort", "quantity_remaining")

        query = """
            SELECT *
            FROM Stock
            WHERE quantity_remaining < ?
        """
        params = [threshold]

        if search:
            query += " AND LOWER(drug_name) LIKE ?"
            params.append(f"%{search}%")

    # -----------------------------
    # Sorting
    # -----------------------------
    allowed_sorts = ["drug_name", "categories", "date", "quantity_remaining"]
    if sort in allowed_sorts:
        query += f" ORDER BY {sort} ASC"
    else:
        query += " ORDER BY quantity_remaining ASC"

    # Execute query
    rows = db_query(query, tuple(params))

    return render_template(
        "stock_low.html",
        user=session.get("user"),
        role=session.get("role"),
        rows=rows,
        search=search,
        sort=sort,
        start_date=start_date,
        end_date=end_date
    )


# ---------------------------------------------------------
# PDF EXPORT FOR LOW STOCK
# ---------------------------------------------------------
@app.route("/stock/low/pdf")
def stock_low_pdf():
    threshold = 10

    rows = db_query("""
        SELECT *
        FROM Stock
        WHERE quantity_remaining <= ?
        ORDER BY quantity_remaining ASC
    """, (threshold,))

    # Build simple HTML content
    html = """
    <html>
    <head>
    <style>
    body { font-family: Arial; }
    h2 { text-align: center; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 1px solid #333; padding: 6px; font-size: 12px; }
    th { background: #eee; }
    </style>
    </head>
    <body>
    <h2>Low Stock Report</h2>
    <table>
    <tr>
        <th>Date</th>
        <th>Drug Name</th>
        <th>Category</th>
        <th>Qty Remaining</th>
        <th>Sales Price</th>
    </tr>
    """

    for r in rows:
        html += f"""
        <tr>
            <td>{r['date']}</td>
            <td>{r['drug_name']}</td>
            <td>{r['categories']}</td>
            <td>{r['quantity_remaining']}</td>
            <td>{r['sales_price']}</td>
        </tr>
        """

    html += "</table></body></html>"

    # Convert HTML to PDF using wkhtmltopdf (Render supports this)
    import pdfkit

    pdf_bytes = pdfkit.from_string(html, False)

    pdf_base64 = base64.b64encode(pdf_bytes).decode("utf-8")

    return jsonify({"pdf": pdf_base64})

# ---------------------------------------------------------
# DRUG LIST
# ---------------------------------------------------------
import os
print("TEMPLATE FOLDER:", app.template_folder)
print("CURRENT WORKING DIR:", os.getcwd())
import os

for root, dirs, files in os.walk(os.getcwd()):
    if "druglist.html" in files:
        print("FOUND:", os.path.join(root, "druglist.html"))

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

    # Permission check
    if not has_permission("drug_entry"):
        return render_template("no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    if request.method == "POST":
        role = session.get("role")

        # Director & Administrator cannot add drugs
        if role in ["Director", "Administrator"]:
            return render_template("no_access.html",
                user=session.get("user"),
                role=session.get("role")
            ), 403

        # Extract form fields
        drug_name = request.form.get("drug_name")
        sales_price = request.form.get("sales_price")
        starting_balance = request.form.get("starting_balance")
        categories = request.form.get("categories")
        cost_price = request.form.get("cost_price")

        # Insert into DrugList table
        db_execute("""
            INSERT INTO DrugList (drug_name, sales_price, starting_balance, categories, cost_price)
            VALUES (?, ?, ?, ?, ?)
        """, (drug_name, sales_price, starting_balance, categories, cost_price))

        # 🔥 Log the drug entry event
        log_event(
            "drug_entry",
            f"Drug entry performed: {drug_name} (Qty: {starting_balance})",
            request.remote_addr,
            session.get("user"),
            session.get("role"),
            request.headers.get("User-Agent")
        )

        return redirect("/druglist")

    return render_template(
        "drug_entry.html",
        user=session.get("user"),
        role=session.get("role")
    )


# ---------------------------------------------------------
# INPATIENT DATA (FULLY UPGRADED)
# ---------------------------------------------------------
@app.route("/inpatient_data")
def inpatient_data():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("inpatient"):
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    rows = db_query("""
        SELECT
            date,
            drug_name,
            patient_name,
            categories,
            sales_price,
            transaction_id,
            bill_total,
            qty_sold,
            invoice_number
        FROM InpatientData
        ORDER BY date DESC, id DESC
    """)

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
import datetime as dt
import qrcode
import io
import base64

# ---------------------------------------------------------
# BILL CALCULATOR PAGE (GET)
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

    # Load all unique patients for autocomplete
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


# ---------------------------------------------------------
# FETCH ALL DATES FOR A PATIENT (AJAX)
# ---------------------------------------------------------
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


# ---------------------------------------------------------
# BILL CALCULATION (GET + POST MERGED)
# ---------------------------------------------------------
@app.route("/inpatient/bill/calculate", methods=["GET", "POST"])
def inpatient_bill_calculate():
    if "user" not in session:
        return redirect("/login")

    if not has_permission("inpatient"):
        return render_template(
            "no_access.html",
            user=session.get("user"),
            role=session.get("role")
        ), 403

    # --------------------------
    # GET → Show calculator page
    # --------------------------
    if request.method == "GET":
        patients = db_query("SELECT DISTINCT patient_name FROM InpatientData ORDER BY patient_name ASC")
        patients = [p["patient_name"] for p in patients]

        return render_template(
            "bill_calculator.html",
            user=session.get("user"),
            role=session.get("role"),
            patients=patients
        )

    # --------------------------
    # POST → Perform calculation
    # --------------------------
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

    generated_date = dt.datetime.now().strftime("%d-%b-%Y %I:%M %p")
    invoice_number = f"INV-{dt.datetime.now().strftime('%Y%m%d%H%M%S')}"

    # Save invoice number to all rows in the range
    db_execute("""
        UPDATE InpatientData
        SET invoice_number = ?
        WHERE LOWER(patient_name) = LOWER(?)
        AND date >= ?
        AND date <= ?
    """, (invoice_number, patient, admission, discharge))

    # --------------------------
    # QR CODE GENERATION
    # --------------------------
    invoice_url = f"http://yourserver.com/invoice/{invoice_number}"

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=6,
        border=2,
    )

    qr.add_data(invoice_url)
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")

    buffer = io.BytesIO()
    img.save(buffer, format="PNG")
    qr_base64 = base64.b64encode(buffer.getvalue()).decode("utf-8")

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
        invoice_number=invoice_number,
        qr_code=qr_base64
    )


# ---------------------------------------------------------
# PATIENT HISTORY PAGE
# ---------------------------------------------------------
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


import qrcode
import io
import base64
from datetime import datetime

@app.route("/inpatient/bill/receipt")
def inpatient_bill_receipt():
    patient = request.args.get("patient")
    admission = request.args.get("admission")
    discharge = request.args.get("discharge")
    invoice_number = request.args.get("invoice")

    # Fetch invoice rows directly by invoice number
    rows = db_query("""
        SELECT date, drug_name, qty_sold, sales_price, bill_total, patient_name
        FROM InpatientData
        WHERE invoice_number = ?
        ORDER BY date ASC
    """, (invoice_number,))

    if not rows:
        return "No inpatient bill found for this invoice number", 404

    # Filter by patient name (case-insensitive)
    rows = [r for r in rows if r["patient_name"].lower() == patient.lower()]
    if not rows:
        return "Patient name does not match invoice records", 404

    # Filter by date range using plain string comparison
    final_rows = [
        r for r in rows
        if admission <= r["date"] <= discharge
    ]

    if not final_rows:
        return "No inpatient bill found for this patient date/range", 404

    total_bill = sum(r["bill_total"] for r in final_rows)
    generated_date = datetime.now().strftime("%d-%b-%Y %I:%M %p")

    base_url = get_public_base_url()
    public_url = f"{base_url}/invoice/{invoice_number}"

    qr = qrcode.make(public_url)
    buffer = io.BytesIO()
    qr.save(buffer, format="PNG")
    qr_base64 = base64.b64encode(buffer.getvalue()).decode("utf-8")

    return render_template(
        "bill_receipt.html",
        patient=patient,
        admission_date=admission,
        discharge_date=discharge,
        bill_items=final_rows,
        total_bill=total_bill,
        generated_date=generated_date,
        invoice_number=invoice_number,
        qr_code=qr_base64,
        public_url=public_url,
        user=session.get("user")
    )

# ================================
# PUBLIC INVOICE VIEW (QR TARGET)
# ================================
@app.route("/invoice/<invoice_number>")
def view_invoice(invoice_number):
    rows = db_query("""
        SELECT date, drug_name, qty_sold, sales_price, bill_total
        FROM InpatientData
        WHERE invoice_number = ?
    """, (invoice_number,))

    if not rows:
        return "Invoice not found", 404

    total_bill = sum(r["bill_total"] for r in rows)
    total_qty = sum(r["qty_sold"] for r in rows)

    # If you generate QR code, insert base64 here
    qr_code = ""

    return render_template(
        "invoice_public.html",
        bill_items=rows,
        total_bill=total_bill,
        total_qty=total_qty,
        invoice_number=invoice_number,
        qr_code=qr_code
    )

def get_public_base_url():
    """
    Automatically detect the correct public-facing base URL.
    Works for localhost, LAN IP, and public domains.
    """
    try:
        # If behind proxy (Nginx, Apache, Render, PythonAnywhere)
        forwarded = request.headers.get("X-Forwarded-Host")
        proto = request.headers.get("X-Forwarded-Proto", "https")

        if forwarded:
            return f"{proto}://{forwarded}"

        # Normal Flask request
        host = request.host
        scheme = request.scheme
        return f"{scheme}://{host}"

    except:
        # Fallback (should never happen)
        return "http://localhost:5000"

# ================================
# PUBLIC VERIFICATION PAGE
# ================================
@app.route("/invoice/verify/<invoice_number>")
def verify_invoice(invoice_number):
    rows = db_query("""
        SELECT date, drug_name, qty_sold, sales_price, bill_total
        FROM InpatientData
        WHERE invoice_number = ?
    """, (invoice_number,))

    if not rows:
        return render_template(
            "invoice_verify.html",
            found=False,
            invoice_number=invoice_number
        )

    total_bill = sum(r["bill_total"] for r in rows)
    total_qty = sum(r["qty_sold"] for r in rows)

    return render_template(
        "invoice_verify.html",
        found=True,
        bill_items=rows,
        total_bill=total_bill,
        total_qty=total_qty,
        invoice_number=invoice_number
    )


# ================================
# DOWNLOADABLE PDF VERSION
# ================================
@app.route("/invoice/<invoice_number>/pdf")
def invoice_pdf(invoice_number):
    rows = db_query("""
        SELECT date, drug_name, qty_sold, sales_price, bill_total
        FROM InpatientData
        WHERE invoice_number = ?
    """, (invoice_number,))

    if not rows:
        return "Invoice not found", 404

    total_bill = sum(r["bill_total"] for r in rows)
    total_qty = sum(r["qty_sold"] for r in rows)

    html = render_template(
        "invoice_pdf.html",
        bill_items=rows,
        total_bill=total_bill,
        total_qty=total_qty,
        invoice_number=invoice_number
    )

    pdf_path = f"invoice_{invoice_number}.pdf"

    # Requires wkhtmltopdf installed
    pdfkit.from_string(html, pdf_path)

    return send_file(
        pdf_path,
        as_attachment=True,
        download_name=f"invoice_{invoice_number}.pdf"
    )

# ---------------------------------------------------------
# DRUG RESTOCK
# ---------------------------------------------------------
@app.route("/drug/restock/save", methods=["POST"])
def drug_restock_save():
    if "user" not in session:
        return redirect("/login")

    role = session.get("role")
    if role in ["Director", "Administrator","HIM Officer"]:
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
    user_id = request.form.get("User ID")
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

    # Check if user exists
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

    # ⭐ LOG EVENT HERE
    log_event(
        "user_create",
        f"New user created: {user_id} ({user_role})",
        request.remote_addr,
        session.get("user"),
        session.get("role"),
        request.headers.get("User-Agent")
    )

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

    # ⭐ LOG EVENT HERE
    log_event(
        "user_deactivate",
        f"User deactivated: {user_id}",
        request.remote_addr,
        session.get("user"),
        session.get("role"),
        request.headers.get("User-Agent")
    )

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

    # Already active?
    if user["Password"] != "DEACTIVATED":
        return redirect("/admin/users?error=User+is+already+active")

    # Reactivate user (set new default password)
    new_password = bcrypt.hashpw("12345".encode("utf-8"), bcrypt.gensalt()).decode("utf-8")

    db_execute("""
        UPDATE Users SET Password=? WHERE [User Id]=?
    """, (new_password, user_id))

    # ⭐ LOG EVENT HERE
    log_event(
        "user_reactivate",
        f"User reactivated: {user_id}",
        request.remote_addr,
        session.get("user"),
        session.get("role"),
        request.headers.get("User-Agent")
    )

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

    # Load the correct logs table
    rows = db_query("SELECT * FROM activity_logs ORDER BY time DESC")

    # Convert rows to dictionaries (prevents blank page)
    logs = [dict(row) for row in rows]

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

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pharmacynew.db")

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

from flask import Flask, request, jsonify
import sqlite3
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pharmacynew.db")



def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


@app.route("/api/kpi_advanced")
def api_kpi_advanced():
    patient_type = request.args.get("patient_type", "All")
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")

    conn = get_db()
    cur = conn.cursor()

    # Stock table structure:
    # id, date, time, drug_name, categories, opening_balance,
    # transaction_type, quantity_sold, quantity_remaining,
    # cost_price, sales_price, customer_type, transaction_id,
    # remark, bill, sales_price_sum, sum_cost_value

    sql = """
        SELECT
            date,
            drug_name,
            categories,
            quantity_sold,
            cost_price,
            sales_price,
            customer_type,
            remark
        FROM Stock
        WHERE 1 = 1
    """
    params = []

    if patient_type != "All":
        sql += " AND customer_type = ?"
        params.append(patient_type)

    if start_date:
        sql += " AND date >= ?"
        params.append(start_date)

    if end_date:
        sql += " AND date <= ?"
        params.append(end_date)

    cur.execute(sql, params)
    rows = cur.fetchall()

    total_sales = 0.0
    total_cost = 0.0
    total_qty = 0.0

    trend = {}               # date -> total sales
    profit_by_drug = {}      # drug -> profit
    qty_by_drug = {}         # drug -> qty
    category_sales = {}      # category -> sales
    patient_bill = {}        # patient_type -> total sales
    expiry_risk = {}         # drug -> qty (if remark suggests expiry)

    # patient_type_trend: { patient_type: { date: sales } }
    patient_type_trend_raw = {}

    for r in rows:
        date = r["date"]
        drug = r["drug_name"]
        cat = r["categories"]
        qty = float(r["quantity_sold"] or 0)
        cost = float(r["cost_price"] or 0)
        sales = float(r["sales_price"] or 0)
        cust = r["customer_type"] or "Unknown"
        remark = r["remark"] or ""

        total_sales += sales
        total_cost += cost
        total_qty += qty

        # overall trend
        trend[date] = trend.get(date, 0) + sales

        # profit by drug
        profit_by_drug[drug] = profit_by_drug.get(drug, 0) + (sales - cost)

        # qty by drug
        qty_by_drug[drug] = qty_by_drug.get(drug, 0) + qty

        # category sales
        category_sales[cat] = category_sales.get(cat, 0) + sales

        # patient type bill
        patient_bill[cust] = patient_bill.get(cust, 0) + sales

        # patient type trend
        if cust not in patient_type_trend_raw:
            patient_type_trend_raw[cust] = {}
        patient_type_trend_raw[cust][date] = patient_type_trend_raw[cust].get(date, 0) + sales

        # expiry risk (simple: remark contains "expiry" or "expire")
        if remark and ("expiry" in remark.lower() or "expire" in remark.lower()):
            expiry_risk[drug] = expiry_risk.get(drug, 0) + qty

    total_profit = total_sales - total_cost
    margin = (total_profit / total_sales * 100) if total_sales > 0 else 0

    def top_dict(d, n=10):
        items = sorted(d.items(), key=lambda x: x[1], reverse=True)
        labels = [k for k, v in items[:n]]
        values = [v for k, v in items[:n]]
        return labels, values

    trend_labels = sorted(trend.keys())
    trend_values = [trend[d] for d in trend_labels]

    top_profit_labels, top_profit_values = top_dict(profit_by_drug, 10)
    top_qty_labels, top_qty_values = top_dict(qty_by_drug, 10)
    category_labels, category_values = top_dict(category_sales, 10)
    patient_type_labels, patient_type_values = top_dict(patient_bill, 10)
    expiry_labels, expiry_values = top_dict(expiry_risk, 10)

    # build patient_type_trend aligned with trend_labels
    patient_type_trend = {}
    for pt in patient_type_labels:
        pt_map = patient_type_trend_raw.get(pt, {})
        patient_type_trend[pt] = [pt_map.get(d, 0) for d in trend_labels]

    response = {
        "total_sales": round(total_sales, 2),
        "total_cost": round(total_cost, 2),
        "total_qty": total_qty,
        "trend_labels": trend_labels,
        "trend_values": trend_values,
        "top_profit_labels": top_profit_labels,
        "top_profit_values": top_profit_values,
        "top_qty_labels": top_qty_labels,
        "top_qty_values": top_qty_values,
        "category_labels": category_labels,
        "category_values": category_values,
        "patient_type_labels": patient_type_labels,
        "patient_type_values": patient_type_values,
        "expiry_labels": expiry_labels,
        "expiry_values": expiry_values,
        "patient_type_trend": patient_type_trend
    }

    conn.close()
    return jsonify(response)
@app.route("/api/kpi_raw")
def api_kpi_raw():
    patient_type = request.args.get("patient_type", "All")
    start_date = request.args.get("start_date")
    end_date = request.args.get("end_date")

    conn = get_db()
    cur = conn.cursor()

    sql = """
        SELECT
            date,
            time,
            drug_name,
            categories,
            quantity_sold,
            quantity_remaining,
            cost_price,
            sales_price,
            customer_type,
            bill,
            remark
        FROM Stock
        WHERE 1 = 1
    """
    params = []

    if patient_type != "All":
        sql += " AND customer_type = ?"
        params.append(patient_type)

    if start_date:
        sql += " AND date >= ?"
        params.append(start_date)

    if end_date:
        sql += " AND date <= ?"
        params.append(end_date)

    cur.execute(sql, params)
    rows = cur.fetchall()

    data = []
    for r in rows:
        data.append({
            "date": r["date"],
            "time": r["time"],
            "drug_name": r["drug_name"],
            "categories": r["categories"],
            "quantity_sold": r["quantity_sold"],
            "quantity_remaining": r["quantity_remaining"],
            "cost_price": r["cost_price"],
            "sales_price": r["sales_price"],
            "customer_type": r["customer_type"],
            "bill": r["bill"],
            "remark": r["remark"],
        })

    conn.close()
    return jsonify(data)

@app.after_request
def add_security_headers(response):
    response.headers["Cache-Control"] = "no-store"
    response.headers["Pragma"] = "no-cache"
    return response

import sqlite3
import user_agents
import requests

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pharmacynew.db")

def get_location(ip):
    try:
        url = f"http://ip-api.com/json/{ip}"
        data = requests.get(url, timeout=2).json()
        return f"{data.get('city')}, {data.get('country')}"
    except:
        return "Unknown"

def log_event(event_type, message, ip, user, role, user_agent_string):
    ua = user_agents.parse(user_agent_string)

    device = f"{ua.device.family} {ua.device.model or ''}".strip()
    browser = f"{ua.browser.family} {ua.browser.version_string}"
    location = get_location(ip)

    conn = sqlite3.connect(DB_PATH, timeout=5)
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO activity_logs (time, type, message, ip, user, role, device, browser, location)
        VALUES (datetime('now', 'localtime'), ?, ?, ?, ?, ?, ?, ?, ?)
    """, (event_type, message, ip, user, role, device, browser, location))

    conn.commit()
    conn.close()

@app.route("/admin/logs/export")
def export_logs():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT * FROM activity_logs ORDER BY time DESC")
    rows = cur.fetchall()
    conn.close()

    csv_data = "Time,Type,Message,IP,User,Role,Device,Browser,Location\n"

    for r in rows:
        csv_data += f"{r['time']},{r['type']},{r['message']},{r['ip']},{r['user']},{r['role']},{r['device']},{r['browser']},{r['location']}\n"

    return csv_data, 200, {
        "Content-Type": "text/csv",
        "Content-Disposition": "attachment; filename=activity_logs.csv"
    }

import time

def safe_execute(query, params=(), retries=3):
    for attempt in range(retries):
        try:
            return db_execute(query, params)
        except sqlite3.OperationalError as e:
            if "locked" in str(e).lower():
                time.sleep(0.2)
            else:
                raise

@app.route("/undo_transaction/<transaction_id>")
def undo_transaction(transaction_id):
    if "user" not in session:
        return redirect("/login")

    # Fetch the transaction
    row = db_query("SELECT * FROM Transactions WHERE id = ?", (transaction_id,))
    if not row:
        return "Transaction not found", 404

    row = row[0]

    # Reverse the stock change
    drug_name = row["DRUG NAME"]
    qty_sold = row["Quantity Sold"]

    db_execute("""
        UPDATE DrugList
        SET [Opening Balance] = [Opening Balance] + ?
        WHERE [DRUG NAME] = ?
    """, (qty_sold, drug_name))

    # Mark transaction as undone (optional)
    db_execute("""
        UPDATE Transactions
        SET status = 'UNDONE'
        WHERE id = ?
    """, (transaction_id,))

    # Log event
    log_event(
        "undo_transaction",
        f"Transaction {transaction_id} undone for drug {drug_name}",
        request.remote_addr,
        session.get("user"),
        session.get("role"),
        request.headers.get("User-Agent")
    )

    return redirect(f"/last_transaction?message=Transaction+{transaction_id}+undone")

@app.route("/last_transaction")
def last_transaction():
    if "user" not in session:
        return redirect("/login")

    transaction_id = session.get("last_transaction_id")
    if not transaction_id:
        return render_template("last_transaction.html", rows=[], transaction_id="N/A")

    rows = db_query("""
        SELECT
            date AS "Date",
            time AS "Time",
            drug_name AS "DRUG NAME",
            categories AS "Category",
            opening_balance AS "Opening Balance",
            transaction_type AS "Transaction Type",
            quantity_sold AS "Quantity Sold",
            quantity_remaining AS "qty Remaining",
            cost_price AS "Cost Price",
            sales_price AS "Sales Price",
            customer_type AS "Customer Type",
            bill AS "Bill",
            sales_price_sum AS "Sales Price Sum",
            sum_cost_value AS "Sum of Cost Price"
        FROM Stock
        WHERE transaction_id = ?
        ORDER BY id ASC
    """, (transaction_id,))

    return render_template(
        "last_transaction.html",
        rows=rows,
        transaction_id=transaction_id
    )

# ---------------------------------------------------------
# RUN APP
# ---------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000, debug=True)
