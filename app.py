import hashlib
import os
import re
import sqlite3
import subprocess
import sys
from datetime import date, datetime, timedelta
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# Streamlit apps must be launched with: python -m streamlit run app.py
if not hasattr(st, "session_state"):
    raise RuntimeError(
        "SHEHaven Bedding requires Streamlit 1.30 or newer. "
        "Install/update with: py -m pip install --upgrade streamlit pandas openpyxl"
    )


# ------------------------------------------------------------
# Direct-run safety
# ------------------------------------------------------------
# Streamlit apps normally need: python -m streamlit run app.py
# This block makes the app friendlier in PyCharm: if someone presses
# the normal green Run button on app.py, it relaunches itself correctly
# through Streamlit.
def _running_inside_streamlit() -> bool:
    if os.environ.get("SHEHAVEN_TEST_MODE") == "1":
        return True
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


if __name__ == "__main__" and not _running_inside_streamlit():
    this_file = Path(__file__).resolve()
    print("Starting SHEHaven Bedding through Streamlit...")
    print(f"Command: {sys.executable} -m streamlit run {this_file}")
    raise SystemExit(subprocess.call([sys.executable, "-m", "streamlit", "run", str(this_file)]))

APP_NAME = "SHEHaven Bedding Stock Management System"
DB_PATH = Path(__file__).with_name("shehaven_bedding.db")

STATUSES = ["Available", "Issued"]
MOVEMENT_TYPES = [
    "Stock In",
    "Sold",
    "Issue Out",
    "Return In",
    "Transfer",
    "Adjustment In",
    "Adjustment Out",
]

# Based on your uploaded SHEHaven Bedding database template.
STOCK_TYPES = [
    "Blankets",
    "Fleece Blankets",
    "Winter Sheets",
    "Fleece Comforters",
    "Velvet Quilts",
    "Throws",
    "Summer Sheets",
    "Duvet Covers",
    "Mattress Protectors",
    "Other",
]

# Brand dropdown list based on the SHEHaven Bedding Excel template you shared.
BRANDS = [
    "No Brand",
    "Mooi mooi",
    "Paris",
    "Pine leaf",
    "Zoya",
    "Orchid",
    "London",
    "Pandora",
    "Electric blankets",
    "Fashion",
    "Pine leaf/Golden",
    "Multi-colored",
    "Plains",
    "Bamboo",
    "Sabana",
    "Other",
]

SIZES = ["Single", "Double", "Queen", "King", "Super King", "Q/K/SK", "One-sized", "1ply", "2ply", "3ply", "Other"]
CONDITIONS = ["New", "Good", "Fair/Worn", "Stained", "Torn", "Damaged", "Missing"]
ROLES = ["Admin", "Manager", "Storekeeper", "Sales User", "Viewer"]

PAYMENT_TYPES = ["Cash Sale", "Credit Sale"]
CREDIT_TERMS = ["1 Month", "2 Months"]
CREDIT_REMINDER_DAYS = 7

EXPENSE_CATEGORIES = [
    "Rent",
    "Transport",
    "Packaging",
    "Wages",
    "Utilities",
    "Marketing",
    "Repairs & Maintenance",
    "Bank Charges",
    "Cleaning",
    "Other",
]

EXPENSE_PAYMENT_METHODS = ["Cash", "EcoCash", "Bank Transfer", "Swipe", "Other"]

SHEHAVEN_TEMPLATE_PRODUCTS = [
    {"stock_type": "Blankets", "brand": "Mooi mooi", "size": "3ply / 2ply"},
    {"stock_type": "Blankets", "brand": "Paris", "size": "1ply / 2ply"},
    {"stock_type": "Blankets", "brand": "Pine leaf", "size": "1ply / 2ply"},
    {"stock_type": "Blankets", "brand": "Zoya", "size": "1ply / 2ply"},
    {"stock_type": "Blankets", "brand": "Orchid", "size": "1ply / 2ply"},
    {"stock_type": "Blankets", "brand": "London", "size": "1ply / 2ply"},
    {"stock_type": "Blankets", "brand": "Pandora", "size": "1ply / 2ply"},
    {"stock_type": "Blankets", "brand": "Electric blankets", "size": "Q/K/SK"},
    {"stock_type": "Fleece Blankets", "brand": "", "size": ""},
    {"stock_type": "Winter Sheets", "brand": "Fashion", "size": "Q/K/SK"},
    {"stock_type": "Fleece Comforters", "brand": "Pine leaf/Golden", "size": "Q/K/SK"},
    {"stock_type": "Velvet Quilts", "brand": "Fashion", "size": "Q/K/SK"},
    {"stock_type": "Throws", "brand": "Multi-colored", "size": "One-sized"},
    {"stock_type": "Throws", "brand": "Plains", "size": "One-sized"},
    {"stock_type": "Summer Sheets", "brand": "Bamboo", "size": "Q/K/SK"},
    {"stock_type": "Summer Sheets", "brand": "Sabana", "size": "Q/K/SK"},
    {"stock_type": "Duvet Covers", "brand": "Fashion", "size": "Q/K/SK"},
    {"stock_type": "Mattress Protectors", "brand": "", "size": "Q/K/SK"},
]

# -----------------------------
# Database helpers
# -----------------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def table_columns(conn, table_name: str) -> set[str]:
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table_name})").fetchall()}


def add_column_if_missing(conn, table_name: str, column_name: str, definition: str):
    if column_name not in table_columns(conn, table_name):
        conn.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {definition}")


def init_db():
    now = datetime.now().isoformat(timespec="seconds")
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS items (
                item_id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_code TEXT UNIQUE NOT NULL,
                stock_type TEXT NOT NULL DEFAULT 'Other',
                item_name TEXT NOT NULL,
                brand TEXT,
                category TEXT NOT NULL DEFAULT 'Bedding',
                size TEXT,
                colour TEXT,
                unit_cost REAL NOT NULL DEFAULT 0,
                selling_price REAL NOT NULL DEFAULT 0,
                supplier TEXT,
                purchase_date TEXT,
                last_restocked_date TEXT,
                reorder_level INTEGER NOT NULL DEFAULT 0,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS locations (
                location_id INTEGER PRIMARY KEY AUTOINCREMENT,
                location_name TEXT UNIQUE NOT NULL,
                location_type TEXT NOT NULL,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS stock_balance (
                balance_id INTEGER PRIMARY KEY AUTOINCREMENT,
                item_id INTEGER NOT NULL,
                location_id INTEGER NOT NULL,
                status TEXT NOT NULL,
                quantity INTEGER NOT NULL DEFAULT 0,
                UNIQUE(item_id, location_id, status),
                FOREIGN KEY(item_id) REFERENCES items(item_id),
                FOREIGN KEY(location_id) REFERENCES locations(location_id)
            );

            CREATE TABLE IF NOT EXISTS stock_movements (
                movement_id INTEGER PRIMARY KEY AUTOINCREMENT,
                movement_date TEXT NOT NULL,
                item_id INTEGER NOT NULL,
                movement_type TEXT NOT NULL,
                quantity INTEGER NOT NULL,
                from_location_id INTEGER,
                from_status TEXT,
                to_location_id INTEGER,
                to_status TEXT,
                customer_name TEXT,
                member_name TEXT,
                mobile_number TEXT,
                sale_amount REAL DEFAULT 0,
                payment_type TEXT DEFAULT 'Cash Sale',
                credit_term TEXT,
                credit_due_date TEXT,
                amount_paid REAL DEFAULT 0,
                balance_due REAL DEFAULT 0,
                payment_status TEXT DEFAULT 'Paid',
                issued_to TEXT,
                approved_by TEXT,
                condition_note TEXT,
                remarks TEXT,
                captured_by TEXT NOT NULL,
                created_at TEXT NOT NULL,
                FOREIGN KEY(item_id) REFERENCES items(item_id),
                FOREIGN KEY(from_location_id) REFERENCES locations(location_id),
                FOREIGN KEY(to_location_id) REFERENCES locations(location_id)
            );


            CREATE TABLE IF NOT EXISTS stocktake (
                stocktake_id INTEGER PRIMARY KEY AUTOINCREMENT,
                count_date TEXT NOT NULL,
                item_id INTEGER NOT NULL,
                location_id INTEGER NOT NULL,
                status TEXT NOT NULL,
                system_quantity INTEGER NOT NULL,
                physical_quantity INTEGER NOT NULL,
                difference INTEGER NOT NULL,
                counted_by TEXT,
                remarks TEXT,
                created_at TEXT NOT NULL,
                FOREIGN KEY(item_id) REFERENCES items(item_id),
                FOREIGN KEY(location_id) REFERENCES locations(location_id)
            );

            CREATE TABLE IF NOT EXISTS expenses (
                expense_id INTEGER PRIMARY KEY AUTOINCREMENT,
                expense_date TEXT NOT NULL,
                category TEXT NOT NULL,
                description TEXT,
                amount REAL NOT NULL DEFAULT 0,
                payment_method TEXT,
                paid_to TEXT,
                receipt_no TEXT,
                captured_by TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            """
        )

        # Safe migrations for users who already created an older SHEHaven database.
        add_column_if_missing(conn, "items", "stock_type", "TEXT NOT NULL DEFAULT 'Other'")
        add_column_if_missing(conn, "items", "brand", "TEXT")
        add_column_if_missing(conn, "items", "selling_price", "REAL DEFAULT 0")
        add_column_if_missing(conn, "items", "last_restocked_date", "TEXT")
        add_column_if_missing(conn, "stock_movements", "customer_name", "TEXT")
        add_column_if_missing(conn, "stock_movements", "member_name", "TEXT")
        add_column_if_missing(conn, "stock_movements", "mobile_number", "TEXT")
        add_column_if_missing(conn, "stock_movements", "sale_amount", "REAL DEFAULT 0")
        add_column_if_missing(conn, "stock_movements", "payment_type", "TEXT DEFAULT 'Cash Sale'")
        add_column_if_missing(conn, "stock_movements", "credit_term", "TEXT")
        add_column_if_missing(conn, "stock_movements", "credit_due_date", "TEXT")
        add_column_if_missing(conn, "stock_movements", "amount_paid", "REAL DEFAULT 0")
        add_column_if_missing(conn, "stock_movements", "balance_due", "REAL DEFAULT 0")
        add_column_if_missing(conn, "stock_movements", "payment_status", "TEXT DEFAULT 'Paid'")
        conn.execute("UPDATE stock_movements SET payment_type = COALESCE(NULLIF(payment_type, ''), 'Cash Sale') WHERE movement_type = 'Sold'")
        conn.execute("UPDATE stock_movements SET amount_paid = COALESCE(amount_paid, sale_amount, 0), balance_due = COALESCE(balance_due, 0), payment_status = COALESCE(NULLIF(payment_status, ''), 'Paid') WHERE movement_type = 'Sold' AND COALESCE(balance_due, 0) <= 0")
        conn.execute("UPDATE items SET stock_type = COALESCE(NULLIF(stock_type, ''), category, 'Other')")
        conn.execute("UPDATE items SET last_restocked_date = COALESCE(last_restocked_date, purchase_date)")

        if conn.execute("SELECT COUNT(*) AS c FROM users WHERE username='admin'").fetchone()["c"] == 0:
            conn.execute(
                """
                INSERT INTO users(full_name, username, password_hash, role, active, created_at)
                VALUES (?, ?, ?, ?, 1, ?)
                """,
                ("SHEHaven Admin", "admin", hash_password("admin123"), "Admin", now),
            )

        for loc_name, loc_type in [
            ("Main Store", "Store"),
            ("Sales Floor", "Store"),
            ("Issued/Client Area", "Issued"),
        ]:
            conn.execute(
                """
                INSERT OR IGNORE INTO locations(location_name, location_type, active, created_at)
                VALUES (?, ?, 1, ?)
                """,
                (loc_name, loc_type, now),
            )
        conn.commit()


def df_query(query: str, params=()):
    with get_conn() as conn:
        return pd.read_sql_query(query, conn, params=params)


def execute(query: str, params=()):
    with get_conn() as conn:
        cur = conn.execute(query, params)
        conn.commit()
        return cur.lastrowid


def authenticate(username: str, password: str):
    with get_conn() as conn:
        user = conn.execute(
            "SELECT * FROM users WHERE username = ? AND active = 1", (username.strip(),)
        ).fetchone()
        if user and user["password_hash"] == hash_password(password):
            return dict(user)
    return None


def slug(text: str, fallback: str = "ITM") -> str:
    cleaned = re.sub(r"[^A-Z0-9]", "", (text or "").upper())
    return (cleaned[:3] or fallback)


def generate_item_code(stock_type: str, brand: str, size: str) -> str:
    prefix = f"{slug(stock_type)}-{slug(brand or size)}"
    with get_conn() as conn:
        count = conn.execute("SELECT COUNT(*) AS c FROM items WHERE item_code LIKE ?", (prefix + "-%",)).fetchone()["c"]
    return f"{prefix}-{count + 1:04d}"


def get_location_id_by_name(name: str) -> int | None:
    with get_conn() as conn:
        row = conn.execute("SELECT location_id FROM locations WHERE location_name = ?", (name,)).fetchone()
        return int(row["location_id"]) if row else None


def item_options():
    df = df_query(
        """
        SELECT
            item_id,
            item_code || ' - ' || COALESCE(stock_type, 'Other') ||
            CASE WHEN COALESCE(brand, '') <> '' THEN ' - ' || brand ELSE '' END ||
            CASE WHEN COALESCE(size, '') <> '' THEN ' (' || size || ')' ELSE '' END AS label
        FROM items
        WHERE active = 1
        ORDER BY stock_type, brand, size, item_code
        """
    )
    return dict(zip(df["label"], df["item_id"])) if not df.empty else {}


def location_options():
    df = df_query(
        """
        SELECT location_id, location_name || ' [' || location_type || ']' AS label
        FROM locations
        WHERE active = 1
        ORDER BY location_name
        """
    )
    return dict(zip(df["label"], df["location_id"])) if not df.empty else {}


def default_select_index(options: list[str], contains: str) -> int:
    for i, option in enumerate(options):
        if contains.lower() in option.lower():
            return i
    return 0


def add_months(original_date: date, months: int) -> date:
    """Add calendar months without needing extra packages."""
    month = original_date.month - 1 + months
    year = original_date.year + month // 12
    month = month % 12 + 1
    month_lengths = [31, 29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    day = min(original_date.day, month_lengths[month - 1])
    return date(year, month, day)


def credit_status_label(due_date_value: str | None, balance_due: float) -> str:
    if float(balance_due or 0) <= 0:
        return "Paid"
    if not due_date_value:
        return "Outstanding"
    try:
        due = datetime.strptime(str(due_date_value), "%Y-%m-%d").date()
    except Exception:
        return "Outstanding"
    today = date.today()
    if due < today:
        return "Overdue"
    if due == today:
        return "Due Today"
    if due <= today + timedelta(days=CREDIT_REMINDER_DAYS):
        return "Due Soon"
    return "Not Yet Due"


def status_quantity(item_id: int, location_id: int, status: str) -> int:
    with get_conn() as conn:
        row = conn.execute(
            """
            SELECT quantity FROM stock_balance
            WHERE item_id = ? AND location_id = ? AND status = ?
            """,
            (item_id, location_id, status),
        ).fetchone()
    return int(row["quantity"]) if row else 0


def total_available_quantity(item_id: int) -> int:
    with get_conn() as conn:
        row = conn.execute(
            """
            SELECT COALESCE(SUM(quantity), 0) AS qty
            FROM stock_balance
            WHERE item_id = ? AND status = 'Available'
            """,
            (item_id,),
        ).fetchone()
    return int(row["qty"])


def get_item_prices(item_id: int) -> tuple[float, float]:
    with get_conn() as conn:
        row = conn.execute(
            "SELECT COALESCE(unit_cost, 0) AS unit_cost, COALESCE(selling_price, 0) AS selling_price FROM items WHERE item_id = ?",
            (item_id,),
        ).fetchone()
    if not row:
        return 0.0, 0.0
    return float(row["unit_cost"] or 0), float(row["selling_price"] or 0)


def change_balance(conn, item_id: int, location_id: int, status: str, delta: int):
    current = conn.execute(
        """
        SELECT quantity FROM stock_balance
        WHERE item_id = ? AND location_id = ? AND status = ?
        """,
        (item_id, location_id, status),
    ).fetchone()
    if current is None:
        if delta < 0:
            raise ValueError("Cannot remove stock from an empty balance.")
        conn.execute(
            """
            INSERT INTO stock_balance(item_id, location_id, status, quantity)
            VALUES (?, ?, ?, ?)
            """,
            (item_id, location_id, status, int(delta)),
        )
    else:
        new_qty = int(current["quantity"]) + int(delta)
        if new_qty < 0:
            raise ValueError(f"Insufficient stock. Current {status} balance at this location is {current['quantity']}.")
        conn.execute(
            """
            UPDATE stock_balance SET quantity = ?
            WHERE item_id = ? AND location_id = ? AND status = ?
            """,
            (new_qty, item_id, location_id, status),
        )


def record_movement(
    movement_date: str,
    item_id: int,
    movement_type: str,
    quantity: int,
    from_location_id: int | None,
    from_status: str | None,
    to_location_id: int | None,
    to_status: str | None,
    issued_to: str = "",
    approved_by: str = "",
    condition_note: str = "Good",
    remarks: str = "",
    captured_by: str = "system",
    customer_name: str = "",
    member_name: str = "",
    mobile_number: str = "",
    sale_amount: float = 0.0,
    payment_type: str = "Cash Sale",
    credit_term: str = "",
    credit_due_date: str = "",
    amount_paid: float = 0.0,
    balance_due: float = 0.0,
    payment_status: str = "Paid",
):
    if quantity <= 0:
        raise ValueError("Quantity must be greater than zero.")
    with get_conn() as conn:
        try:
            conn.execute("BEGIN")
            if from_location_id and from_status:
                change_balance(conn, item_id, from_location_id, from_status, -quantity)
            if to_location_id and to_status and movement_type != "Sold":
                change_balance(conn, item_id, to_location_id, to_status, quantity)

            movement_id = conn.execute(
                """
                INSERT INTO stock_movements(
                    movement_date, item_id, movement_type, quantity, from_location_id, from_status,
                    to_location_id, to_status, customer_name, member_name, mobile_number, sale_amount, payment_type, credit_term,
                    credit_due_date, amount_paid, balance_due, payment_status, issued_to, approved_by,
                    condition_note, remarks, captured_by, created_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    movement_date,
                    item_id,
                    movement_type,
                    int(quantity),
                    from_location_id,
                    from_status,
                    to_location_id,
                    to_status,
                    customer_name,
                    member_name,
                    mobile_number,
                    float(sale_amount or 0),
                    payment_type or "Cash Sale",
                    credit_term or "",
                    credit_due_date or "",
                    float(amount_paid or 0),
                    float(balance_due or 0),
                    payment_status or "Paid",
                    issued_to,
                    approved_by,
                    condition_note,
                    remarks,
                    captured_by,
                    datetime.now().isoformat(timespec="seconds"),
                ),
            ).lastrowid

            if movement_type == "Stock In":
                conn.execute(
                    "UPDATE items SET last_restocked_date = ? WHERE item_id = ?",
                    (movement_date, item_id),
                )

            conn.commit()
            return movement_id
        except Exception:
            conn.rollback()
            raise


def create_item(
    item_code: str,
    stock_type: str,
    brand: str,
    size: str,
    colour: str,
    unit_cost: float,
    selling_price: float,
    purchase_date: str,
    last_restocked_date: str,
    reorder_level: int,
    opening_qty: int = 0,
    location_id: int | None = None,
):
    item_name = " - ".join([part for part in [stock_type, brand, size] if part])
    item_code = item_code.strip().upper() or generate_item_code(stock_type, brand, size)
    now = datetime.now().isoformat(timespec="seconds")
    with get_conn() as conn:
        try:
            conn.execute("BEGIN")
            item_id = conn.execute(
                """
                INSERT INTO items(
                    item_code, stock_type, item_name, brand, category, size, colour, unit_cost, selling_price,
                    supplier, purchase_date, last_restocked_date, reorder_level, active, created_at
                ) VALUES (?, ?, ?, ?, 'Bedding', ?, ?, ?, ?, '', ?, ?, ?, 1, ?)
                """,
                (
                    item_code,
                    stock_type,
                    item_name,
                    brand,
                    size,
                    colour,
                    float(unit_cost or 0),
                    float(selling_price or 0),
                    purchase_date,
                    last_restocked_date,
                    int(reorder_level or 0),
                    now,
                ),
            ).lastrowid
            if opening_qty > 0:
                if location_id is None:
                    location_id = get_location_id_by_name("Main Store")
                conn.execute(
                    """
                    INSERT INTO stock_balance(item_id, location_id, status, quantity)
                    VALUES (?, ?, 'Available', ?)
                    """,
                    (item_id, location_id, int(opening_qty)),
                )
                conn.execute(
                    """
                    INSERT INTO stock_movements(
                        movement_date, item_id, movement_type, quantity, to_location_id, to_status,
                        condition_note, remarks, captured_by, created_at
                    ) VALUES (?, ?, 'Stock In', ?, ?, 'Available', 'New', 'Opening stock captured when item was created', ?, ?)
                    """,
                    (last_restocked_date, item_id, int(opening_qty), location_id, st.session_state.get("user", {}).get("username", "system"), now),
                )
            conn.commit()
            return item_id, item_code
        except Exception:
            conn.rollback()
            raise


def current_stock_df():
    return df_query(
        """
        SELECT
            i.item_code AS "Item Code",
            i.stock_type AS "Stock Type",
            i.brand AS "Brand",
            i.size AS "Size",
            i.colour AS "Colours",
            l.location_name AS "Location",
            l.location_type AS "Location Type",
            b.status AS "Status",
            b.quantity AS "Quantity",
            i.unit_cost AS "Cost Price",
            i.selling_price AS "Selling Price",
            b.quantity * i.unit_cost AS "Stock Value",
            b.quantity * i.selling_price AS "Potential Sales Value",
            b.quantity * (i.selling_price - i.unit_cost) AS "Potential Gross Profit",
            i.reorder_level AS "Reorder Level",
            i.last_restocked_date AS "Last Restocked Date"
        FROM stock_balance b
        JOIN items i ON i.item_id = b.item_id
        JOIN locations l ON l.location_id = b.location_id
        WHERE b.quantity > 0 AND i.active = 1 AND l.active = 1
        ORDER BY i.stock_type, i.brand, i.size, l.location_name, b.status
        """
    )


def stock_summary_df():
    stock = current_stock_df()
    if stock.empty:
        return pd.DataFrame(columns=["Item Code", "Stock Type", "Brand", "Size", "Colours", "Total", "Available", "Issued", "Reorder Level", "Last Restocked Date"])
    pivot = stock.pivot_table(
        index=["Item Code", "Stock Type", "Brand", "Size", "Colours", "Reorder Level", "Last Restocked Date"],
        columns="Status",
        values="Quantity",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    for status in STATUSES:
        if status not in pivot.columns:
            pivot[status] = 0
    pivot["Total"] = pivot[STATUSES].sum(axis=1)
    cols = ["Item Code", "Stock Type", "Brand", "Size", "Colours", "Total", "Available", "Issued", "Reorder Level", "Last Restocked Date"]
    return pivot[cols].sort_values(["Stock Type", "Brand", "Size"])


def movements_df():
    return df_query(
        """
        SELECT
            m.movement_id AS "Movement ID",
            m.movement_date AS "Date",
            i.item_code AS "Item Code",
            i.stock_type AS "Stock Type",
            i.brand AS "Brand",
            i.size AS "Size",
            m.movement_type AS "Movement Type",
            m.quantity AS "Quantity",
            fl.location_name AS "From Location",
            m.from_status AS "From Status",
            tl.location_name AS "To Location",
            m.to_status AS "To Status",
            m.customer_name AS "Customer / Receipt No.",
            m.member_name AS "Member Name",
            m.mobile_number AS "Mobile Number",
            m.sale_amount AS "Sale Amount",
            m.payment_type AS "Payment Type",
            m.credit_term AS "Credit Term",
            m.credit_due_date AS "Credit Due Date",
            m.amount_paid AS "Amount Paid",
            m.balance_due AS "Balance Due",
            m.payment_status AS "Payment Status",
            m.issued_to AS "Issued To / Received By",
            m.approved_by AS "Approved By",
            m.condition_note AS "Condition",
            m.remarks AS "Remarks",
            m.captured_by AS "Captured By",
            m.created_at AS "Captured At"
        FROM stock_movements m
        JOIN items i ON i.item_id = m.item_id
        LEFT JOIN locations fl ON fl.location_id = m.from_location_id
        LEFT JOIN locations tl ON tl.location_id = m.to_location_id
        ORDER BY m.movement_date DESC, m.movement_id DESC
        """
    )


def sold_stock_df():
    mv = movements_df()
    if mv.empty:
        return mv
    return mv[mv["Movement Type"] == "Sold"]


def credit_sales_df():
    df = df_query(
        """
        SELECT
            m.movement_id AS "Sale ID",
            m.movement_date AS "Sale Date",
            i.item_code AS "Item Code",
            i.stock_type AS "Stock Type",
            i.brand AS "Brand",
            i.size AS "Size",
            m.customer_name AS "Customer / Receipt No.",
            m.member_name AS "Member Name",
            m.mobile_number AS "Mobile Number",
            m.quantity AS "Quantity",
            m.sale_amount AS "Sale Amount",
            m.amount_paid AS "Amount Paid",
            m.balance_due AS "Balance Due",
            m.credit_term AS "Credit Term",
            m.credit_due_date AS "Credit Due Date",
            m.payment_status AS "Payment Status",
            m.remarks AS "Remarks",
            m.captured_by AS "Captured By",
            m.created_at AS "Captured At"
        FROM stock_movements m
        JOIN items i ON i.item_id = m.item_id
        WHERE m.movement_type = 'Sold'
          AND COALESCE(m.payment_type, 'Cash Sale') = 'Credit Sale'
        ORDER BY m.credit_due_date ASC, m.movement_id DESC
        """
    )
    if df.empty:
        return df
    df["Reminder Status"] = df.apply(lambda r: credit_status_label(r.get("Credit Due Date"), r.get("Balance Due")), axis=1)
    return df


def credit_reminders_df():
    credits = credit_sales_df()
    if credits.empty:
        return credits
    return credits[credits["Reminder Status"].isin(["Overdue", "Due Today", "Due Soon", "Outstanding"])]


def update_credit_payment(sale_id: int, amount_received: float, remarks: str, captured_by: str):
    if amount_received <= 0:
        raise ValueError("Amount received must be greater than zero.")
    with get_conn() as conn:
        row = conn.execute(
            """
            SELECT sale_amount, amount_paid, balance_due, payment_status
            FROM stock_movements
            WHERE movement_id = ? AND movement_type = 'Sold'
            """,
            (sale_id,),
        ).fetchone()
        if not row:
            raise ValueError("Sale record was not found.")
        new_paid = float(row["amount_paid"] or 0) + float(amount_received)
        sale_amount = float(row["sale_amount"] or 0)
        new_balance = max(sale_amount - new_paid, 0.0)
        new_status = "Paid" if new_balance <= 0 else "Part Paid"
        existing_remarks = ""
        conn.execute(
            """
            UPDATE stock_movements
            SET amount_paid = ?, balance_due = ?, payment_status = ?, remarks = COALESCE(remarks, '') || ?
            WHERE movement_id = ?
            """,
            (
                new_paid,
                new_balance,
                new_status,
                f"\nPayment received {date.today().isoformat()}: ${amount_received:,.2f} by {captured_by}. {remarks}".rstrip(),
                sale_id,
            ),
        )
        conn.commit()
        return new_balance


def products_df():
    """Product register shown in the app and reports. Supplier is intentionally excluded."""
    return df_query(
        """
        SELECT
            item_id AS "Product ID",
            item_code AS "Item Code",
            stock_type AS "Stock Type",
            brand AS "Brand",
            size AS "Size",
            colour AS "Colours",
            unit_cost AS "Cost Price",
            selling_price AS "Selling Price",
            reorder_level AS "Reorder Level",
            last_restocked_date AS "Last Restocked Date",
            active AS "Active",
            created_at AS "Created At"
        FROM items
        ORDER BY stock_type, brand, size, item_code
        """
    )


def stocktake_df():
    """Stocktake history for reporting."""
    return df_query(
        """
        SELECT
            s.stocktake_id AS "Stocktake ID",
            s.count_date AS "Count Date",
            i.item_code AS "Item Code",
            i.stock_type AS "Stock Type",
            i.brand AS "Brand",
            i.size AS "Size",
            l.location_name AS "Location",
            s.status AS "Status",
            s.system_quantity AS "System Quantity",
            s.physical_quantity AS "Physical Quantity",
            s.difference AS "Difference",
            s.counted_by AS "Counted By",
            s.remarks AS "Remarks",
            s.created_at AS "Captured At"
        FROM stocktake s
        JOIN items i ON i.item_id = s.item_id
        JOIN locations l ON l.location_id = s.location_id
        ORDER BY s.count_date DESC, s.stocktake_id DESC
        """
    )


def expenses_df():
    """Expense register for running costs such as rent, transport, wages, packaging, and utilities."""
    return df_query(
        """
        SELECT
            expense_id AS "Expense ID",
            expense_date AS "Expense Date",
            category AS "Category",
            description AS "Description",
            amount AS "Amount",
            payment_method AS "Payment Method",
            paid_to AS "Paid To",
            receipt_no AS "Receipt / Reference",
            captured_by AS "Captured By",
            created_at AS "Captured At"
        FROM expenses
        ORDER BY expense_date DESC, expense_id DESC
        """
    )


def sales_profit_df():
    """Sales report with estimated cost of goods sold and gross profit."""
    return df_query(
        """
        SELECT
            m.movement_id AS "Sale ID",
            m.movement_date AS "Sale Date",
            i.item_code AS "Item Code",
            i.stock_type AS "Stock Type",
            i.brand AS "Brand",
            i.size AS "Size",
            m.customer_name AS "Customer / Receipt No.",
            m.member_name AS "Member Name",
            m.mobile_number AS "Mobile Number",
            m.quantity AS "Quantity",
            i.unit_cost AS "Cost Price",
            CASE WHEN m.quantity > 0 THEN m.sale_amount / m.quantity ELSE 0 END AS "Selling Price Used",
            m.sale_amount AS "Sale Amount",
            COALESCE(m.amount_paid, 0) AS "Amount Collected",
            COALESCE(m.balance_due, 0) AS "Balance Due",
            COALESCE(m.payment_type, 'Cash Sale') AS "Payment Type",
            COALESCE(m.payment_status, 'Paid') AS "Payment Status",
            m.quantity * COALESCE(i.unit_cost, 0) AS "Cost of Goods Sold",
            m.sale_amount - (m.quantity * COALESCE(i.unit_cost, 0)) AS "Estimated Gross Profit",
            m.captured_by AS "Captured By",
            m.created_at AS "Captured At"
        FROM stock_movements m
        JOIN items i ON i.item_id = m.item_id
        WHERE m.movement_type = 'Sold'
        ORDER BY m.movement_date DESC, m.movement_id DESC
        """
    )


def profit_summary_df():
    """One-line business summary combining sales, credit balances, and expenses."""
    sales = sales_profit_df()
    expenses = expenses_df()
    total_sales = float(sales["Sale Amount"].sum()) if not sales.empty else 0.0
    amount_collected = float(sales["Amount Collected"].sum()) if not sales.empty else 0.0
    outstanding_credit = float(sales["Balance Due"].sum()) if not sales.empty else 0.0
    cogs = float(sales["Cost of Goods Sold"].sum()) if not sales.empty else 0.0
    gross_profit = float(sales["Estimated Gross Profit"].sum()) if not sales.empty else 0.0
    total_expenses = float(expenses["Amount"].sum()) if not expenses.empty else 0.0
    net_profit = gross_profit - total_expenses
    return pd.DataFrame([
        {
            "Total Sales": total_sales,
            "Amount Collected": amount_collected,
            "Outstanding Credit": outstanding_credit,
            "Cost of Goods Sold": cogs,
            "Estimated Gross Profit": gross_profit,
            "Total Expenses": total_expenses,
            "Estimated Net Profit After Expenses": net_profit,
        }
    ])

def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, df in sheets.items():
            clean_name = name[:31]
            df.to_excel(writer, sheet_name=clean_name, index=False)
            ws = writer.sheets[clean_name]
            for col in ws.columns:
                max_length = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_length + 2, 42)
    return output.getvalue()

# -----------------------------
# UI helpers
# -----------------------------
def login_screen():
    st.title("🛏️ SHEHaven Bedding")
    st.subheader("Stock Management System")
    st.info("Default login: username `admin`, password `admin123`.")
    with st.form("login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
        if submitted:
            user = authenticate(username, password)
            if user:
                st.session_state["user"] = user
                st.rerun()
            else:
                st.error("Invalid username or password.")


def require_roles(*roles):
    user = st.session_state.get("user", {})
    return user.get("role") in roles


def sidebar():
    user = st.session_state.get("user")
    st.sidebar.title("SHEHaven 🛏️")
    st.sidebar.caption(f"Logged in as **{user['full_name']}** · {user['role']}")
    pages = [
        "Dashboard",
        "Product Register",
        "Add Stock",
        "Sales",
        "Expenses",
        "Credit Reminders",
        "Stocktake",
        "Reports",
    ]
    if user["role"] == "Admin":
        pages.append("User Management")
    page = st.sidebar.radio("Navigate", pages)
    if st.sidebar.button("Logout"):
        st.session_state.clear()
        st.rerun()
    st.sidebar.divider()
    st.sidebar.caption("Every sheet accounted for.")
    return page


def show_df(df: pd.DataFrame, height=420):
    if df.empty:
        st.info("No records found yet.")
    else:
        st.dataframe(df, use_container_width=True, height=height)


def load_shehaven_template_products():
    inserted = 0
    skipped = 0
    main_store = get_location_id_by_name("Main Store")
    for item in SHEHAVEN_TEMPLATE_PRODUCTS:
        code = generate_item_code(item["stock_type"], item["brand"], item["size"])
        try:
            create_item(
                item_code=code,
                stock_type=item["stock_type"],
                brand=item["brand"],
                size=item["size"],
                colour="",
                unit_cost=0,
                selling_price=0,
                purchase_date=date.today().isoformat(),
                last_restocked_date=date.today().isoformat(),
                reorder_level=0,
                opening_qty=0,
                location_id=main_store,
            )
            inserted += 1
        except sqlite3.IntegrityError:
            skipped += 1
    return inserted, skipped

# -----------------------------
# Pages
# -----------------------------
def page_dashboard():
    st.title("📊 SHEHaven Bedding Dashboard")
    stock = current_stock_df()
    summary = stock_summary_df()
    movements = movements_df()

    total_items = int(stock["Quantity"].sum()) if not stock.empty else 0
    total_value = float(stock["Stock Value"].sum()) if not stock.empty else 0
    available = int(stock.loc[stock["Status"] == "Available", "Quantity"].sum()) if not stock.empty else 0
    issued = int(stock.loc[stock["Status"] == "Issued", "Quantity"].sum()) if not stock.empty else 0
    sold_today = 0
    sales_today = 0.0
    total_sold_qty = 0
    total_sales_value = 0.0
    cash_sales_value = 0.0
    credit_sales_value = 0.0
    if not movements.empty:
        all_sold = movements[movements["Movement Type"] == "Sold"]
        if not all_sold.empty:
            total_sold_qty = int(all_sold["Quantity"].sum())
            total_sales_value = float(all_sold["Sale Amount"].sum())
            cash_sales_value = float(all_sold.loc[all_sold["Payment Type"].fillna("Cash Sale") == "Cash Sale", "Sale Amount"].sum())
            credit_sales_value = float(all_sold.loc[all_sold["Payment Type"].fillna("Cash Sale") == "Credit Sale", "Sale Amount"].sum())
        today_sold = all_sold[all_sold["Date"] == date.today().isoformat()] if not all_sold.empty else pd.DataFrame()
        sold_today = int(today_sold["Quantity"].sum()) if not today_sold.empty else 0
        sales_today = float(today_sold["Sale Amount"].sum()) if not today_sold.empty else 0.0

    reminders = credit_reminders_df()
    expenses = expenses_df()
    sales_profit = sales_profit_df()
    outstanding_credit = credit_sales_df()
    outstanding_amount = float(outstanding_credit["Balance Due"].sum()) if not outstanding_credit.empty else 0.0
    reminders_count = len(reminders.index) if not reminders.empty else 0
    expenses_today = 0.0
    month_expenses = 0.0
    month_gross_profit = 0.0
    month_net_profit = 0.0
    current_month = date.today().strftime("%Y-%m")
    if not expenses.empty:
        expenses_today = float(expenses.loc[expenses["Expense Date"] == date.today().isoformat(), "Amount"].sum())
        month_expenses = float(expenses.loc[expenses["Expense Date"].astype(str).str.startswith(current_month), "Amount"].sum())
    if not sales_profit.empty:
        month_sales = sales_profit[sales_profit["Sale Date"].astype(str).str.startswith(current_month)]
        month_gross_profit = float(month_sales["Estimated Gross Profit"].sum()) if not month_sales.empty else 0.0
    month_net_profit = month_gross_profit - month_expenses

    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Total Stock", f"{total_items:,}")
    c2.metric("Available", f"{available:,}")
    c3.metric("Total Sales Qty", f"{total_sold_qty:,}")
    c4.metric("Total Sales Value", f"${total_sales_value:,.2f}")
    c5.metric("Outstanding Credit", f"${outstanding_amount:,.2f}")
    c6.metric("Stock Value", f"${total_value:,.2f}")

    s1, s2, s3, s4 = st.columns(4)
    s1.metric("Sold Today", f"{sold_today:,}")
    s2.metric("Sales Today", f"${sales_today:,.2f}")
    s3.metric("Cash Sales Value", f"${cash_sales_value:,.2f}")
    s4.metric("Credit Sales Value", f"${credit_sales_value:,.2f}")

    e1, e2, e3 = st.columns(3)
    e1.metric("Expenses Today", f"${expenses_today:,.2f}")
    e2.metric("Expenses This Month", f"${month_expenses:,.2f}")
    e3.metric("Est. Net Profit This Month", f"${month_net_profit:,.2f}")

    if sales_today > 0:
        st.success(f"Today’s recorded sales value: ${sales_today:,.2f}")
    if reminders_count > 0:
        st.warning(f"Credit reminder: {reminders_count} credit sale(s) are overdue, due today, or due within {CREDIT_REMINDER_DAYS} days.")
        show_df(reminders[["Sale ID", "Sale Date", "Member Name", "Mobile Number", "Sale Amount", "Amount Paid", "Balance Due", "Credit Term", "Credit Due Date", "Reminder Status"]], height=220)

    if not summary.empty:
        low_stock = summary[summary["Available"] <= summary["Reorder Level"]]
        if not low_stock.empty:
            st.warning("Some items are at or below reorder level.")
            show_df(low_stock[["Item Code", "Stock Type", "Brand", "Size", "Available", "Reorder Level"]], height=200)

    left, right = st.columns(2)
    with left:
        st.subheader("Stock by Status")
        if not stock.empty:
            st.bar_chart(stock.groupby("Status")["Quantity"].sum())
        else:
            st.info("Add products and stock to populate the dashboard.")
    with right:
        st.subheader("Stock by Type")
        if not stock.empty:
            st.bar_chart(stock.groupby("Stock Type")["Quantity"].sum())
        else:
            st.info("No stock type data yet.")

    st.subheader("Current Stock Summary")
    show_df(summary, height=360)

    if st.button("Load SHEHaven product list from template"):
        inserted, skipped = load_shehaven_template_products()
        st.success(f"Template products loaded. Added {inserted}, skipped {skipped}.")
        st.rerun()


def page_product_register():
    st.title("🧾 Product Register")
    tab1, tab2, tab3 = st.tabs(["Add Product", "Add Location", "View Registers"])

    with tab1:
        if not require_roles("Admin", "Manager", "Storekeeper"):
            st.error("You do not have permission to add products.")
        else:
            with st.form("add_product_form"):
                st.caption("Use Stock Type the same way your Excel file uses Product name.")
                col1, col2, col3 = st.columns(3)
                stock_type = col1.selectbox("Stock Type / Product Name", STOCK_TYPES)
                brand_choice = col2.selectbox("Brand", BRANDS)
                size = col3.text_input("Size", placeholder="e.g. Q/K/SK, 1ply / 2ply")
                if brand_choice == "Other":
                    brand = st.text_input("Enter New Brand Name", placeholder="Type the brand name")
                elif brand_choice == "No Brand":
                    brand = ""
                else:
                    brand = brand_choice

                col4, col5, col6, col7 = st.columns(4)
                colours = col4.text_input("Colours")
                unit_cost = col5.number_input("Cost Price", min_value=0.0, step=0.01, format="%.2f")
                selling_price = col6.number_input("Selling Price", min_value=0.0, step=0.01, format="%.2f")
                reorder_level = col7.number_input("Reorder Level", min_value=0, step=1, value=0)

                col8, col9 = st.columns(2)
                current_qty = col8.number_input("Current Stock Qty", min_value=0, step=1, value=0)
                last_restocked_date = col9.date_input("Last Restocked Date", value=date.today())

                item_code = st.text_input("Item Code", placeholder="Leave blank to auto-generate, e.g. BLA-PAR-0001")
                submitted = st.form_submit_button("Save Product")
                if submitted:
                    try:
                        _, final_code = create_item(
                            item_code=item_code,
                            stock_type=stock_type,
                            brand=brand.strip(),
                            size=size.strip(),
                            colour=colours.strip(),
                            unit_cost=float(unit_cost),
                            selling_price=float(selling_price),
                            purchase_date=last_restocked_date.isoformat(),
                            last_restocked_date=last_restocked_date.isoformat(),
                            reorder_level=int(reorder_level),
                            opening_qty=int(current_qty),
                            location_id=get_location_id_by_name("Main Store"),
                        )
                        st.success(f"Product saved successfully. Item Code: {final_code}")
                    except sqlite3.IntegrityError:
                        st.error("That item code already exists. Use another code or leave it blank for auto-generation.")
                    except Exception as e:
                        st.error(f"Could not save product: {e}")

    with tab2:
        if not require_roles("Admin", "Manager"):
            st.error("You do not have permission to add locations.")
        else:
            with st.form("add_location_form"):
                col1, col2 = st.columns(2)
                location_name = col1.text_input("Location Name", placeholder="e.g. Branch A Store")
                location_type = col2.selectbox("Location Type", ["Store", "Sales Floor", "Branch", "Other"])
                submitted = st.form_submit_button("Save Location")
                if submitted:
                    if not location_name.strip():
                        st.error("Location name is required.")
                    else:
                        try:
                            execute(
                                """
                                INSERT INTO locations(location_name, location_type, active, created_at)
                                VALUES (?, ?, 1, ?)
                                """,
                                (location_name.strip(), location_type, datetime.now().isoformat(timespec="seconds")),
                            )
                            st.success("Location saved successfully.")
                        except sqlite3.IntegrityError:
                            st.error("That location already exists.")

    with tab3:
        st.subheader("Products")
        show_df(products_df(), height=320)
        st.subheader("Locations")
        locations = df_query(
            "SELECT location_name AS 'Location', location_type AS 'Type', active AS 'Active' FROM locations ORDER BY location_name"
        )
        show_df(locations, height=220)


def page_stock_control():
    st.title("📦 Stock Control")
    if not require_roles("Admin", "Manager", "Storekeeper", "Sales User"):
        st.error("You do not have permission to record stock changes.")
        return

    item_map = item_options()
    loc_map = location_options()
    if not item_map or not loc_map:
        st.info("Add products and locations first.")
        return

    tab1, tab2, tab3 = st.tabs(["➕ Add Stock", "🧾 Remove Sold Stock", "🔁 Advanced Movement"])

    with tab1:
        st.subheader("Add Stock")
        st.caption("Use this when new stock arrives or when you restock existing products.")
        with st.form("quick_add_stock_form"):
            col1, col2, col3 = st.columns(3)
            movement_date = col1.date_input("Date", value=date.today(), key="add_stock_date")
            item_label = col2.selectbox("Product", list(item_map.keys()), key="add_stock_item")
            quantity = col3.number_input("Quantity to Add", min_value=1, step=1, value=1)
            loc_labels = list(loc_map.keys())
            col4, col5, col6 = st.columns(3)
            to_label = col4.selectbox("Add To Location", loc_labels, index=default_select_index(loc_labels, "Main Store"))
            condition_note = col5.selectbox("Condition", CONDITIONS, index=0)
            approved_by = col6.text_input("Approved By")
            remarks = st.text_area("Remarks", placeholder="e.g. New delivery received")
            submitted = st.form_submit_button("➕ Add Stock")
            if submitted:
                try:
                    record_movement(
                        movement_date.isoformat(),
                        item_map[item_label],
                        "Stock In",
                        int(quantity),
                        None,
                        None,
                        loc_map[to_label],
                        "Available",
                        approved_by=approved_by.strip(),
                        condition_note=condition_note,
                        remarks=remarks.strip(),
                        captured_by=st.session_state["user"]["username"],
                    )
                    st.success("Stock added successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not add stock: {e}")

    with tab2:
        st.subheader("Remove Sold Stock")
        st.caption("Use this button when an item is sold. It removes quantity from Available stock and records the sale in movement history.")
        with st.form("remove_sold_stock_form"):
            col1, col2, col3 = st.columns(3)
            sale_date = col1.date_input("Sale Date", value=date.today(), key="sold_stock_date")
            item_label = col2.selectbox("Product Sold", list(item_map.keys()), key="sold_stock_item")
            quantity = col3.number_input("Quantity Sold", min_value=1, step=1, value=1)
            available_qty = total_available_quantity(item_map[item_label])
            st.info(f"Total available stock for this product: {available_qty:,}")
            loc_labels = list(loc_map.keys())
            _, default_selling_price = get_item_prices(item_map[item_label])
            col4, col5, col6 = st.columns(3)
            from_label = col4.selectbox("Remove From Location", loc_labels, index=default_select_index(loc_labels, "Main Store"))
            customer = col5.text_input("Customer / Receipt No.")
            unit_selling_price = col6.number_input("Selling Price per Unit", min_value=0.0, value=float(default_selling_price), step=0.01, format="%.2f")
            payment_type = st.selectbox("Payment Type", PAYMENT_TYPES, key="stock_control_payment_type")
            credit_term = ""
            credit_due_date = ""
            amount_paid = 0.0
            member_name = ""
            mobile_number = ""
            sale_amount = float(unit_selling_price) * int(quantity)
            if payment_type == "Credit Sale":
                cmem, cmob = st.columns(2)
                member_name = cmem.text_input("Member Name", key="stock_control_member_name")
                mobile_number = cmob.text_input("Mobile Number", key="stock_control_mobile_number", placeholder="e.g. 0771234567")
                cterm, cdeposit = st.columns(2)
                credit_term = cterm.selectbox("Credit Type", CREDIT_TERMS, key="stock_control_credit_term")
                months = 1 if credit_term == "1 Month" else 2
                credit_due_date = add_months(sale_date, months).isoformat()
                amount_paid = cdeposit.number_input("Deposit / Amount Paid Now", min_value=0.0, max_value=float(sale_amount), value=0.0, step=0.01, format="%.2f")
                st.warning(f"Credit payment due date: {credit_due_date}")
            else:
                amount_paid = sale_amount
            balance_due = max(float(sale_amount) - float(amount_paid), 0.0)
            st.success(f"Sale total to be recorded: ${sale_amount:,.2f}")
            if payment_type == "Credit Sale":
                st.info(f"Balance due: ${balance_due:,.2f}")
            remarks = st.text_area("Remarks", placeholder="e.g. Cash sale / credit sale / receipt number")
            submitted = st.form_submit_button("🧾 Remove Sold Stock")
            if submitted:
                if payment_type == "Credit Sale" and (not member_name.strip() or not mobile_number.strip()):
                    st.error("For credit sales, please enter both Member Name and Mobile Number.")
                else:
                    try:
                        record_movement(
                            sale_date.isoformat(),
                            item_map[item_label],
                            "Sold",
                            int(quantity),
                            loc_map[from_label],
                            "Available",
                            None,
                            None,
                            issued_to=customer.strip(),
                            customer_name=customer.strip(),
                            member_name=member_name.strip(),
                            mobile_number=mobile_number.strip(),
                            sale_amount=float(sale_amount),
                            payment_type=payment_type,
                            credit_term=credit_term,
                            credit_due_date=credit_due_date,
                            amount_paid=float(amount_paid),
                            balance_due=float(balance_due),
                            payment_status="Paid" if balance_due <= 0 else ("Part Paid" if amount_paid > 0 else "Outstanding"),
                            condition_note="Good",
                            remarks=remarks.strip(),
                            captured_by=st.session_state["user"]["username"],
                        )
                        st.success("Sold stock removed successfully.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
                    except Exception as e:
                        st.error(f"Could not remove sold stock: {e}")

    with tab3:
        st.subheader("Advanced Stock Movement")
        with st.form("movement_form"):
            col1, col2, col3 = st.columns(3)
            movement_date = col1.date_input("Movement Date", value=date.today())
            movement_type = col2.selectbox("Movement Type", MOVEMENT_TYPES)
            item_label = col3.selectbox("Product", list(item_map.keys()))
            quantity = st.number_input("Quantity", min_value=1, step=1, value=1)
            st.caption("Use source fields when stock must leave a location. Use destination fields when stock must enter a location/status.")
            loc_labels = list(loc_map.keys())
            col4, col5 = st.columns(2)
            from_label = col4.selectbox("From Location", ["None"] + loc_labels)
            from_status = col5.selectbox("From Status", ["None"] + STATUSES)
            col6, col7 = st.columns(2)
            to_label = col6.selectbox("To Location", ["None"] + loc_labels)
            to_status = col7.selectbox("To Status", ["None"] + STATUSES)
            col8, col9, col10 = st.columns(3)
            issued_to = col8.text_input("Issued To / Received By")
            approved_by = col9.text_input("Approved By")
            condition_note = col10.selectbox("Condition", CONDITIONS)
            remarks = st.text_area("Remarks")
            submitted = st.form_submit_button("Record Movement")
            if submitted:
                item_id = item_map[item_label]
                from_location_id = None if from_label == "None" else loc_map[from_label]
                to_location_id = None if to_label == "None" else loc_map[to_label]
                from_status_value = None if from_status == "None" else from_status
                to_status_value = None if to_status == "None" else to_status
                if not from_location_id and not to_location_id:
                    st.error("Select at least a source or destination location.")
                elif (from_location_id and not from_status_value) or (to_location_id and not to_status_value and movement_type != "Sold"):
                    st.error("Select the relevant status for each selected location.")
                else:
                    try:
                        record_movement(
                            movement_date.isoformat(),
                            item_id,
                            movement_type,
                            int(quantity),
                            from_location_id,
                            from_status_value,
                            to_location_id,
                            to_status_value,
                            issued_to.strip(),
                            approved_by.strip(),
                            condition_note,
                            remarks.strip(),
                            st.session_state["user"]["username"],
                        )
                        st.success("Stock movement recorded successfully.")
                        st.rerun()
                    except ValueError as e:
                        st.error(str(e))
                    except Exception as e:
                        st.error(f"Could not record movement: {e}")

    st.subheader("Recent Movements")
    show_df(movements_df().head(100), height=420)


def page_add_stock():
    st.title("➕ Add Stock")
    st.caption("Use this section when new stock arrives or when you want to increase stock for an existing product.")

    if not require_roles("Admin", "Manager", "Storekeeper"):
        st.error("You do not have permission to add stock.")
        return

    item_map = item_options()
    loc_map = location_options()
    if not item_map or not loc_map:
        st.info("Add products and locations first. Go to Product Register, then come back here to add stock.")
        return

    with st.form("dedicated_add_stock_form"):
        col1, col2, col3 = st.columns(3)
        movement_date = col1.date_input("Stock Date", value=date.today(), key="dedicated_add_stock_date")
        item_label = col2.selectbox("Product", list(item_map.keys()), key="dedicated_add_stock_item")
        quantity = col3.number_input("Quantity to Add", min_value=1, step=1, value=1)

        loc_labels = list(loc_map.keys())
        col4, col5, col6 = st.columns(3)
        to_label = col4.selectbox("Add To Location", loc_labels, index=default_select_index(loc_labels, "Main Store"))
        condition_note = col5.selectbox("Condition", CONDITIONS, index=0)
        approved_by = col6.text_input("Received / Approved By")

        remarks = st.text_area("Remarks", placeholder="Example: New delivery received / restocked from supplier")
        submitted = st.form_submit_button("➕ Add Stock to System")

        if submitted:
            try:
                record_movement(
                    movement_date.isoformat(),
                    item_map[item_label],
                    "Stock In",
                    int(quantity),
                    None,
                    None,
                    loc_map[to_label],
                    "Available",
                    approved_by=approved_by.strip(),
                    condition_note=condition_note,
                    remarks=remarks.strip(),
                    captured_by=st.session_state["user"]["username"],
                )
                st.success("Stock added successfully.")
                st.rerun()
            except Exception as e:
                st.error(f"Could not add stock: {e}")

    st.subheader("Recent Stock Added")
    added = movements_df()
    if not added.empty:
        added = added[added["Movement Type"] == "Stock In"].head(50)
    show_df(added, height=360)


def page_sales():
    st.title("🧾 Sales")
    st.caption("Use this section when stock is sold. Track cash sales and credit sales, including 1-month or 2-month payment terms.")

    if not require_roles("Admin", "Manager", "Sales User"):
        st.error("You do not have permission to record sales.")
        return

    item_map = item_options()
    loc_map = location_options()
    if not item_map or not loc_map:
        st.info("Add products and stock first before recording sales.")
        return

    with st.form("dedicated_sales_form"):
        col1, col2, col3 = st.columns(3)
        sale_date = col1.date_input("Sale Date", value=date.today(), key="dedicated_sale_date")
        item_label = col2.selectbox("Product Sold", list(item_map.keys()), key="dedicated_sale_item")
        quantity = col3.number_input("Quantity Sold", min_value=1, step=1, value=1)

        item_id = item_map[item_label]
        available_qty = total_available_quantity(item_id)
        unit_cost, default_selling_price = get_item_prices(item_id)
        st.info(f"Total available stock for this product: {available_qty:,}")

        loc_labels = list(loc_map.keys())
        col4, col5, col6 = st.columns(3)
        from_label = col4.selectbox("Sold From Location", loc_labels, index=default_select_index(loc_labels, "Main Store"))
        customer = col5.text_input("Customer / Receipt No.")
        unit_selling_price = col6.number_input(
            "Selling Price per Unit",
            min_value=0.0,
            value=float(default_selling_price),
            step=0.01,
            format="%.2f",
        )

        sale_amount = float(unit_selling_price) * int(quantity)
        gross_profit = (float(unit_selling_price) - float(unit_cost)) * int(quantity)

        col7, col8, col9 = st.columns(3)
        payment_type = col7.selectbox("Payment Type", PAYMENT_TYPES, key="dedicated_payment_type")
        credit_term = ""
        credit_due_date = ""
        member_name = ""
        mobile_number = ""
        if payment_type == "Credit Sale":
            mcol1, mcol2 = st.columns(2)
            member_name = mcol1.text_input("Member Name", key="dedicated_member_name")
            mobile_number = mcol2.text_input("Mobile Number", key="dedicated_mobile_number", placeholder="e.g. 0771234567")
            credit_term = col8.selectbox("Credit Type", CREDIT_TERMS, key="dedicated_credit_term")
            months = 1 if credit_term == "1 Month" else 2
            credit_due_date = add_months(sale_date, months).isoformat()
            amount_paid = col9.number_input("Deposit / Amount Paid Now", min_value=0.0, max_value=float(sale_amount), value=0.0, step=0.01, format="%.2f")
        else:
            col8.text_input("Credit Type", value="Not applicable", disabled=True)
            col9.text_input("Credit Due Date", value="Not applicable", disabled=True)
            amount_paid = sale_amount

        balance_due = max(float(sale_amount) - float(amount_paid), 0.0)
        payment_status = "Paid" if balance_due <= 0 else ("Part Paid" if amount_paid > 0 else "Outstanding")

        c1, c2, c3 = st.columns(3)
        c1.success(f"Sale total: ${sale_amount:,.2f}")
        c2.info(f"Estimated gross profit: ${gross_profit:,.2f}")
        if payment_type == "Credit Sale":
            c3.warning(f"Due {credit_due_date} · Balance ${balance_due:,.2f}")
        else:
            c3.success("Cash sale: paid now")

        remarks = st.text_area("Remarks", placeholder="Example: Cash sale / EcoCash / credit sale / invoice number")
        submitted = st.form_submit_button("🧾 Record Sale and Remove Stock")

        if submitted:
            if int(quantity) > int(available_qty):
                st.error(f"Cannot sell {quantity:,}. Only {available_qty:,} available in total.")
            elif payment_type == "Credit Sale" and (not member_name.strip() or not mobile_number.strip()):
                st.error("For credit sales, please enter both Member Name and Mobile Number.")
            else:
                try:
                    record_movement(
                        sale_date.isoformat(),
                        item_id,
                        "Sold",
                        int(quantity),
                        loc_map[from_label],
                        "Available",
                        None,
                        None,
                        issued_to=customer.strip(),
                        customer_name=customer.strip(),
                        member_name=member_name.strip(),
                        mobile_number=mobile_number.strip(),
                        sale_amount=float(sale_amount),
                        payment_type=payment_type,
                        credit_term=credit_term,
                        credit_due_date=credit_due_date,
                        amount_paid=float(amount_paid),
                        balance_due=float(balance_due),
                        payment_status=payment_status,
                        condition_note="Good",
                        remarks=remarks.strip(),
                        captured_by=st.session_state["user"]["username"],
                    )
                    st.success("Sale recorded and stock removed successfully.")
                    st.rerun()
                except ValueError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"Could not record sale: {e}")

    st.subheader("Recent Sales")
    sales = sold_stock_df()
    show_df(sales.head(100) if not sales.empty else sales, height=420)



def page_expenses():
    st.title("💸 Expenses")
    st.caption("Use this section to record business expenses so SHEHaven can compare sales, costs, and profit.")

    if not require_roles("Admin", "Manager"):
        st.error("You do not have permission to capture or view expenses.")
        return

    with st.form("expense_form"):
        col1, col2, col3 = st.columns(3)
        expense_date = col1.date_input("Expense Date", value=date.today())
        category = col2.selectbox("Expense Category", EXPENSE_CATEGORIES)
        amount = col3.number_input("Amount", min_value=0.0, step=0.01, format="%.2f")

        col4, col5, col6 = st.columns(3)
        payment_method = col4.selectbox("Payment Method", EXPENSE_PAYMENT_METHODS)
        paid_to = col5.text_input("Paid To", placeholder="Person / company paid")
        receipt_no = col6.text_input("Receipt / Reference No.", placeholder="Receipt, invoice or EcoCash ref")

        description = st.text_area("Description / Notes", placeholder="Example: Transport for new stock delivery")
        submitted = st.form_submit_button("💸 Save Expense")
        if submitted:
            if amount <= 0:
                st.error("Expense amount must be greater than zero.")
            else:
                try:
                    execute(
                        """
                        INSERT INTO expenses(expense_date, category, description, amount, payment_method, paid_to, receipt_no, captured_by, created_at)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            expense_date.isoformat(),
                            category,
                            description.strip(),
                            float(amount),
                            payment_method,
                            paid_to.strip(),
                            receipt_no.strip(),
                            st.session_state["user"]["username"],
                            datetime.now().isoformat(timespec="seconds"),
                        ),
                    )
                    st.success("Expense saved successfully.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not save expense: {e}")

    expenses = expenses_df()
    sales_profit = sales_profit_df()

    total_expenses = float(expenses["Amount"].sum()) if not expenses.empty else 0.0
    today_expenses = float(expenses.loc[expenses["Expense Date"] == date.today().isoformat(), "Amount"].sum()) if not expenses.empty else 0.0
    current_month = date.today().strftime("%Y-%m")
    month_expenses = float(expenses.loc[expenses["Expense Date"].astype(str).str.startswith(current_month), "Amount"].sum()) if not expenses.empty else 0.0
    month_gross_profit = 0.0
    if not sales_profit.empty:
        month_sales = sales_profit[sales_profit["Sale Date"].astype(str).str.startswith(current_month)]
        month_gross_profit = float(month_sales["Estimated Gross Profit"].sum()) if not month_sales.empty else 0.0
    month_net_profit = month_gross_profit - month_expenses

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Expenses", f"${total_expenses:,.2f}")
    c2.metric("Today", f"${today_expenses:,.2f}")
    c3.metric("This Month", f"${month_expenses:,.2f}")
    c4.metric("Est. Net Profit This Month", f"${month_net_profit:,.2f}")

    if not expenses.empty:
        st.subheader("Expenses by Category")
        st.bar_chart(expenses.groupby("Category")["Amount"].sum().sort_values(ascending=False))

    st.subheader("Recent Expenses")
    show_df(expenses.head(100) if not expenses.empty else expenses, height=420)

def page_credit_reminders():
    st.title("⏰ Credit Reminders")
    st.caption("This section shows credit sales that are overdue, due today, or due within the next 7 days.")

    if not require_roles("Admin", "Manager", "Sales User"):
        st.error("You do not have permission to view credit reminders.")
        return

    credits = credit_sales_df()
    reminders = credit_reminders_df()

    if credits.empty:
        st.info("No credit sales recorded yet.")
        return

    total_credit = float(credits["Sale Amount"].sum())
    total_paid = float(credits["Amount Paid"].sum())
    total_balance = float(credits["Balance Due"].sum())
    overdue = credits[credits["Reminder Status"] == "Overdue"]
    due_soon = credits[credits["Reminder Status"].isin(["Due Today", "Due Soon"])]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Credit Sales", f"${total_credit:,.2f}")
    c2.metric("Amount Paid", f"${total_paid:,.2f}")
    c3.metric("Balance Due", f"${total_balance:,.2f}")
    c4.metric("Needs Reminder", f"{len(reminders.index):,}")

    if not overdue.empty:
        st.error("Overdue credit payments")
        show_df(overdue[["Sale ID", "Sale Date", "Member Name", "Mobile Number", "Balance Due", "Credit Due Date", "Reminder Status"]], height=220)
    if not due_soon.empty:
        st.warning("Credit payments due soon")
        show_df(due_soon[["Sale ID", "Sale Date", "Member Name", "Mobile Number", "Balance Due", "Credit Due Date", "Reminder Status"]], height=220)

    st.subheader("Record Credit Payment")
    outstanding = credits[credits["Balance Due"] > 0].copy()
    if outstanding.empty:
        st.success("All credit sales are fully paid.")
    else:
        outstanding["Payment Label"] = outstanding.apply(
            lambda r: f"Sale {int(r['Sale ID'])} - {r['Member Name']} ({r['Mobile Number']}) - Balance ${float(r['Balance Due']):,.2f} - Due {r['Credit Due Date']}",
            axis=1,
        )
        sale_map = dict(zip(outstanding["Payment Label"], outstanding["Sale ID"]))
        with st.form("record_credit_payment_form"):
            sale_label = st.selectbox("Credit Sale", list(sale_map.keys()))
            amount_received = st.number_input("Amount Received", min_value=0.0, step=0.01, format="%.2f")
            payment_remarks = st.text_area("Payment Remarks", placeholder="Example: Paid cash / EcoCash reference / bank transfer")
            submitted = st.form_submit_button("Mark Payment Received")
            if submitted:
                try:
                    new_balance = update_credit_payment(
                        int(sale_map[sale_label]),
                        float(amount_received),
                        payment_remarks.strip(),
                        st.session_state["user"]["username"],
                    )
                    if new_balance <= 0:
                        st.success("Payment recorded. Credit sale is now fully paid.")
                    else:
                        st.success(f"Payment recorded. New balance: ${new_balance:,.2f}")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not record payment: {e}")

    st.subheader("All Credit Sales")
    show_df(credits, height=460)


def page_stocktake():
    st.title("📦 Stocktake")
    if not require_roles("Admin", "Manager", "Storekeeper"):
        st.error("You do not have permission to capture stocktake records.")
        return
    item_map = item_options()
    loc_map = location_options()
    if not item_map or not loc_map:
        st.info("Add products and locations first.")
        return
    with st.form("stocktake_form"):
        col1, col2, col3 = st.columns(3)
        count_date = col1.date_input("Count Date", value=date.today())
        item_label = col2.selectbox("Product", list(item_map.keys()))
        location_label = col3.selectbox("Location", list(loc_map.keys()))
        col4, col5 = st.columns(2)
        status = col4.selectbox("Status", STATUSES)
        physical_quantity = col5.number_input("Physical Quantity", min_value=0, step=1, value=0)
        counted_by = st.text_input("Counted By", value=st.session_state["user"]["full_name"])
        remarks = st.text_area("Remarks")
        submitted = st.form_submit_button("Save Stocktake")
        if submitted:
            item_id = item_map[item_label]
            location_id = loc_map[location_label]
            system_quantity = status_quantity(item_id, location_id, status)
            difference = int(physical_quantity) - int(system_quantity)
            execute(
                """
                INSERT INTO stocktake(count_date, item_id, location_id, status, system_quantity, physical_quantity, difference, counted_by, remarks, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    count_date.isoformat(),
                    item_id,
                    location_id,
                    status,
                    system_quantity,
                    int(physical_quantity),
                    difference,
                    counted_by.strip(),
                    remarks.strip(),
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
            st.success(f"Stocktake saved. Difference: {difference:+d}")
            if difference != 0:
                st.warning("A difference was found. Record an Adjustment In or Adjustment Out if management approves the correction.")
    st.subheader("Stocktake History")
    show_df(stocktake_df(), height=400)


def page_reports():
    st.title("📑 Reports")
    stock = current_stock_df()
    summary = stock_summary_df()
    movements = movements_df()
    stocktakes = stocktake_df()
    products = products_df()
    sold = sold_stock_df()
    credit_sales = credit_sales_df()
    credit_reminders = credit_reminders_df()
    expenses = expenses_df()
    sales_profit = sales_profit_df()
    profit_summary = profit_summary_df()

    report = st.selectbox(
        "Choose Report",
        ["Products", "Current Stock", "Stock Summary", "Stock Movements", "Sold Stock", "Sales Profit", "Credit Sales", "Credit Reminders", "Expenses", "Profit Summary", "Low Stock", "Stocktake"],
    )
    if report == "Products":
        df = products
    elif report == "Current Stock":
        df = stock
    elif report == "Stock Summary":
        df = summary
    elif report == "Stock Movements":
        df = movements
    elif report == "Sold Stock":
        df = sold
    elif report == "Sales Profit":
        df = sales_profit
    elif report == "Credit Sales":
        df = credit_sales
    elif report == "Credit Reminders":
        df = credit_reminders
    elif report == "Expenses":
        df = expenses
    elif report == "Profit Summary":
        df = profit_summary
    elif report == "Low Stock":
        df = summary[summary["Available"] <= summary["Reorder Level"]] if not summary.empty else pd.DataFrame()
    else:
        df = stocktakes

    show_df(df, height=500)
    st.download_button(
        "Download selected report as CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name=f"SHEHaven_{report.replace(' ', '_')}_{date.today().isoformat()}.csv",
        mime="text/csv",
        disabled=df.empty,
    )
    excel_bytes = to_excel_bytes(
        {
            "Products": products,
            "Current Stock": stock,
            "Stock Summary": summary,
            "Movements": movements,
            "Sold Stock": sold,
            "Sales Profit": sales_profit,
            "Credit Sales": credit_sales,
            "Credit Reminders": credit_reminders,
            "Expenses": expenses,
            "Profit Summary": profit_summary,
            "Stocktake": stocktakes,
        }
    )
    st.download_button(
        "Download full workbook as Excel",
        data=excel_bytes,
        file_name=f"SHEHaven_Bedding_Reports_{date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def page_user_management():
    st.title("👥 User Management")
    if not require_roles("Admin"):
        st.error("Admin access required.")
        return
    with st.form("add_user_form"):
        col1, col2, col3 = st.columns(3)
        full_name = col1.text_input("Full Name")
        username = col2.text_input("Username")
        role = col3.selectbox("Role", ROLES)
        password = st.text_input("Temporary Password", type="password")
        submitted = st.form_submit_button("Create User")
        if submitted:
            if not full_name.strip() or not username.strip() or not password:
                st.error("Full name, username, and password are required.")
            else:
                try:
                    execute(
                        """
                        INSERT INTO users(full_name, username, password_hash, role, active, created_at)
                        VALUES (?, ?, ?, ?, 1, ?)
                        """,
                        (
                            full_name.strip(),
                            username.strip(),
                            hash_password(password),
                            role,
                            datetime.now().isoformat(timespec="seconds"),
                        ),
                    )
                    st.success("User created successfully.")
                except sqlite3.IntegrityError:
                    st.error("That username already exists.")
    st.subheader("Users")
    users = df_query(
        "SELECT full_name AS 'Full Name', username AS 'Username', role AS 'Role', active AS 'Active', created_at AS 'Created At' FROM users ORDER BY full_name"
    )
    show_df(users, height=350)


def main():
    st.set_page_config(page_title="SHEHaven Bedding", page_icon="🛏️", layout="wide")
    init_db()
    if "user" not in st.session_state:
        login_screen()
        return
    page = sidebar()
    if page == "Dashboard":
        page_dashboard()
    elif page == "Product Register":
        page_product_register()
    elif page == "Add Stock":
        page_add_stock()
    elif page == "Sales":
        page_sales()
    elif page == "Expenses":
        page_expenses()
    elif page == "Credit Reminders":
        page_credit_reminders()
    elif page == "Stocktake":
        page_stocktake()
    elif page == "Reports":
        page_reports()
    elif page == "User Management":
        page_user_management()


if __name__ == "__main__":
    main()
