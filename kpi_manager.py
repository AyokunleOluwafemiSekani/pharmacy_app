import sqlite3, os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pharmacy.db")

def get_kpis():
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    # Sales today
    cur.execute("SELECT SUM(total_amount) FROM sales WHERE date = DATE('now')")
    sales_today = cur.fetchone()[0] or 0

    # Sales this month
    cur.execute("SELECT SUM(total_amount) FROM sales WHERE strftime('%Y-%m', date) = strftime('%Y-%m', 'now')")
    sales_month = cur.fetchone()[0] or 0

    # Total transactions
    cur.execute("SELECT COUNT(*) FROM sales")
    total_transactions = cur.fetchone()[0]

    # Total customers
    cur.execute("SELECT COUNT(*) FROM customers")
    total_customers = cur.fetchone()[0]

    # Top selling products
    cur.execute("""
        SELECT product_name, SUM(quantity) AS qty 
        FROM sales_items 
        GROUP BY product_name 
        ORDER BY qty DESC 
        LIMIT 5
    """)
    top_products = cur.fetchall()

    # Low stock alerts
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
import json, os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SETTINGS_FILE = os.path.join(BASE_DIR, "kpi_settings.json")

def load_kpi_settings():
    if not os.path.exists(SETTINGS_FILE):
        return {
            "show_sales_today": True,
            "show_sales_month": True,
            "show_transactions": True,
            "show_customers": True,
            "show_top_products": True,
            "show_low_stock": True,
            "low_stock_threshold": 10,
            "date_range": "monthly"
        }
    return json.load(open(SETTINGS_FILE))

def save_kpi_settings(data):
    json.dump(data, open(SETTINGS_FILE, "w"), indent=4)
