from flask import Flask, render_template, request, redirect, url_for, session
from io import BytesIO
from flask import send_file
from werkzeug.security import check_password_hash
from datetime import datetime
import pandas as pd
import os
import sys
import gspread

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

from google.oauth2.service_account import Credentials
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from functools import lru_cache

app = Flask(
    __name__,
    template_folder=resource_path("templates"),
    static_folder=resource_path("static")
)
app.secret_key = "vms_secret_key"
app.permanent_session_lifetime = 120   # 2 minutes

# GOOGLE SHEETS CONFIG

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

creds_path = resource_path("credentials.json")

creds = Credentials.from_service_account_file(
    creds_path,
    scopes=SCOPES
)

gs_client = gspread.authorize(creds)

gs_book = gs_client.open_by_key(
    "1By6HFnyJTwL_0tLcW0BwM74QpsPKvrQyFvheomBS-p0"
)

# =========================
# USER SYSTEM
# =========================

def get_users():
    return read_sheet("users")

def require_login():
    if "user" not in session:
        return redirect(url_for("home"))
    return None
    login_check = require_login()
    if login_check:
        return login_check

# ==============================
# VEHICLE SUMMARY DATA ENGINE
# ==============================

@lru_cache(maxsize=32)
def read_sheet(sheet_name):
    print("Reading from Google Sheets:", sheet_name)   # debug line
    ws = gs_book.worksheet(sheet_name)
    data = ws.get_all_records()
    return pd.DataFrame(data)

def filter_vehicle(df, vehicle_no):
    if "Vehicle No" in df.columns:
        return df[df["Vehicle No"].astype(str) == str(vehicle_no)]
    return df


def pick_columns(df, cols):
    exist = [c for c in cols if c in df.columns]
    return df[exist]


def latest_rows(df, count=5, date_col="Date"):
    if date_col in df.columns:
        df = df.sort_values(date_col, ascending=False)
    return df.head(count)

# =========================
# PATH CONFIG
# =========================

BASE_DIR = r"C:\excel_web_project"
EXCEL_FILE = os.path.join(BASE_DIR, "data", "vehicle_data.xlsx")

# =========================
# FILTER MAP
# =========================

FILTERS = {
    "basic": ["Vehicle No", "Vehicle Type","Purpose"],
    "service": ["Vehicle No", "Service Done By"],
    "repair": ["Vehicle No", "Service Done By"],
    "add_work": ["Vehicle No", "Service Done By"],
    "pending": ["Vehicle No", "Pending Works"],
    "expenditure": ["Vehicle No", "Service Done By"],
    "tyre": ["Vehicle No", "Purchase By"],
    "fuel": ["Vehicle No"],
    "revenue": ["Vehicle No", "Expair Month"],
    "exp_sum": ["Month"],
    "fuel_sum": ["Month"],
    "prov_bal": []
}

# =========================
# SAFE LOAD + FORMAT
# =========================

def load_sheet(sheet):

    df = read_sheet(sheet)

    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):

            def fmt(x):
                if pd.isna(x):
                    return ""
                try:
                    x = float(x)
                    if x.is_integer():
                        return f"{int(x):,}"
                    else:
                        return f"{x:,.2f}"
                except:
                    return x

            df[col] = df[col].apply(fmt)

    return df


def apply_filters(df, key):
    for col in FILTERS.get(key, []):
        val = request.args.get(col)
        if val and val != "ALL" and col in df.columns:
            df = df[df[col].astype(str) == val]
    return df


def get_filter_values(sheet, key):
    df = read_sheet(sheet)
    out = {}
    for col in FILTERS.get(key, []):
        if col in df.columns:
            out[col] = sorted(df[col].dropna().astype(str).unique())
    return out


# =========================
# AUTH ROUTES
# =========================

@app.route("/")
def home():
    if "user" in session:
        return redirect(url_for("report_page"))
    return render_template("index.html")

@app.route("/login", methods=["POST"])
def login():

    username = request.form.get("username")
    password = request.form.get("password")

    users_df = get_users()

    if "username" not in users_df.columns:
        return render_template("index.html",
                               error="Users sheet not configured properly")

    user_row = users_df[
        users_df["username"].astype(str) == str(username)
    ]

    if user_row.empty:
        return render_template("index.html",
                               error="Invalid Username or Password")

    user = user_row.iloc[0]

    if str(user.get("status", "")).strip() != "Active":
        return render_template("index.html",
                               error="User is Inactive")

    if not check_password_hash(
            str(user.get("password_hash", "")),
            password
        ):
        return render_template("index.html",
                               error="Invalid Username or Password")

    # Store session data
    session["user"] = user["username"]
    session["role"] = user.get("role", "")
    session["assigned_vehicles"] = user.get("assigned_vehicles", "")
    session["allowed_reports"] = user.get("allowed_reports", "")

    return redirect(url_for("report_page"))

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("home"))

# =========================
# REPORT PAGE
# =========================

@app.route("/report")
def report_page():

    login_check = require_login()
    if login_check:
        return login_check

    return render_template("report.html",
                           username=session["user"],
                           role=session["role"])

# ==============================
# VEHICLE SUMMARY PAGE ROUTE
# ==============================

@app.route("/vehicle/<vehicle_no>")
def vehicle_summary(vehicle_no):

    if require_login():
        return require_login()

    role = session.get("role", "")
    assigned = session.get("assigned_vehicles", "")

    if role in ["Driver", "Executive"]:
        allowed = [v.strip() for v in assigned.split(",") if v.strip()]
        if vehicle_no not in allowed:
            return "Access Denied", 403

    # ---------- BASIC DETAILS ----------
    basic_cols = [
        "Purpose","To Whom","Designation","Vehicle Type","Make","Model","Chassis Number",
        "Engine No","Type Of Fuel","Year Registered","Purpose",
        "Valuation Date","Valuation Value"
    ]

    basic_df = read_sheet("Basic Details")
    basic_df = filter_vehicle(basic_df, vehicle_no)
    basic_df = pick_columns(basic_df, basic_cols)
    basic = basic_df.head(1).to_dict("records")


    # ---------- REPAIR ----------
    repair_cols = ["Date","Milage","Description","Service Done By"]
    repair_df = read_sheet("Repair")
    repair_df = filter_vehicle(repair_df, vehicle_no)
    repair_df = pick_columns(repair_df, repair_cols)
    repair_df = latest_rows(repair_df, 10)
    repair = repair_df.to_dict("records")


    # ---------- SERVICE ----------
    service_cols = [
        "Date","Milage","Description","Service Done By",
        "Next Service Mailage","Next Sevice Date"
    ]
    service_df = read_sheet("Service")
    service_df = filter_vehicle(service_df, vehicle_no)
    service_df = pick_columns(service_df, service_cols)
    service_df = latest_rows(service_df, 5)
    service = service_df.to_dict("records")


    # ---------- ADD WORK ----------
    add_cols = ["Date","Milage","Description","Service Done By"]
    add_df = read_sheet("add_work")
    add_df = filter_vehicle(add_df, vehicle_no)
    add_df = pick_columns(add_df, add_cols)
    add_df = latest_rows(add_df, 5)
    add_work = add_df.to_dict("records")


    # ---------- PENDING ----------
    pending_cols = ["Status","Pending Works","Pending Works 1","Remarks"]
    pending_df = read_sheet("Pending_Works")
    pending_df = filter_vehicle(pending_df, vehicle_no)
    pending_df = pick_columns(pending_df, pending_cols)
    pending = pending_df.head(1).to_dict("records")


    # ---------- TYRE ----------
    tyre_cols = [
        "Date","Milage","Model/Country","No of Tyre",
        "Purchase By","Next Tyre Milage"
    ]
    tyre_df = read_sheet("Tyre1")
    tyre_df = filter_vehicle(tyre_df, vehicle_no)
    tyre_df = pick_columns(tyre_df, tyre_cols)
    tyre_df = latest_rows(tyre_df, 5)
    tyre = tyre_df.to_dict("records")


    # ---------- FUEL ----------
    fuel_cols = [
        "2026-February","2026-January","2025-December",
        "2025-November","2025-October","2025-September"
    ]
    fuel_df = read_sheet("Fuel Cost")
    fuel_df = filter_vehicle(fuel_df, vehicle_no)
    fuel_df = pick_columns(fuel_df, fuel_cols)
    fuel = fuel_df.head(1).to_dict("records")


    # ---------- REVNUE ----------
    rev_cols = ["Expair Month","Expair Date"]
    rev_df = read_sheet("Revnue Li")
    rev_df = filter_vehicle(rev_df, vehicle_no)
    rev_df = pick_columns(rev_df, rev_cols)
    revnue = rev_df.head(1).to_dict("records")


   # ---------- IMAGE ----------
    vehicle_image = f"vehicle_images/{vehicle_no}.jpg"
    return render_template(
    	"vehicle_summary.html",
    	vehicle_no=vehicle_no,
    	vehicle_image=vehicle_image,
    	basic=basic,
    	repair=repair,
    	service=service,
   	add_work=add_work,
   	pending=pending,
    	tyre=tyre,
    	fuel=fuel,
    	revnue=revnue
    )

# =========================
# GENERIC REPORT VIEW
# =========================

def render_report(sheet, title, key):

    if require_login():
        return require_login()

    role = session.get("role", "")
    allowed_reports = session.get("allowed_reports", "")

    if role == "Auditor":
        allowed = [r.strip().lower() for r in allowed_reports.split(",") if r.strip()]
        if key.lower() not in allowed:
            return "Access Denied", 403

    df = load_sheet(sheet)
    df = apply_filters(df, key)

    html_table = df.to_html(
        classes="table table-bordered table-striped",
        index=False,
        justify="right"
    )

    return render_template(
        "table.html",
        title=title,
        table=html_table,
        filters=get_filter_values(sheet, key),
        report_key=key
    )

# =========================
# PDF ENGINE (RAM VERSION)
# =========================

def render_pdf(sheet, title, key):

    df = load_sheet(sheet)
    df = apply_filters(df, key)

    buffer = BytesIO()   # <-- Create memory buffer

    doc = SimpleDocTemplate(
        buffer,                  # <-- IMPORTANT (not filename)
        pagesize=landscape(A4)
    )

    styles = getSampleStyleSheet()
    elements = []

    elements.append(Paragraph(title, styles['Title']))
    elements.append(Spacer(1, 12))

    if df.empty:
        elements.append(Paragraph("No data available.", styles['Normal']))
    else:
        data = [df.columns.tolist()] + df.values.tolist()

        table = Table(data, repeatRows=1)

        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
            ('FONTSIZE', (0, 0), (-1, -1), 8)
        ]))

        elements.append(table)

    doc.build(elements)

    buffer.seek(0)   # <-- VERY IMPORTANT

    return buffer

# =========================
# REPORT ROUTES (12)
# =========================

@app.route("/report/basic")
def basic(): return render_report("Basic Details", "Basic Details", "basic")

@app.route("/report/service")
def service(): return render_report("Service", "Service", "service")

@app.route("/report/repair")
def repair(): return render_report("Repair", "Repair", "repair")

@app.route("/report/add_work")
def add_work(): return render_report("add_work", "add work", "add_work")

@app.route("/report/pending")
def pending(): return render_report("Pending_Works", "Pending Works", "pending")

@app.route("/report/expenditure")
def expenditure(): return render_report("Expenditure", "Expenditure", "expenditure")

@app.route("/report/tyre")
def tyre(): return render_report("Tyre1", "Tyre Replace", "tyre")

@app.route("/report/fuel")
def fuel(): return render_report("Fuel Cost", "Fuel Cost", "fuel")

@app.route("/report/revenue")
def revenue(): return render_report("Revnue Li", "Revenue License", "revenue")

@app.route("/report/exp_sum")
def exp_sum(): return render_report("exp_sum", "Expenditure Summary", "exp_sum")

@app.route("/report/fuel_sum")
def fuel_sum(): return render_report("Fuel Sum", "Fuel Summary", "fuel_sum")

@app.route("/report/prov_bal")
def prov_bal(): return render_report("prov_bal", "Provision Balance", "prov_bal")


# =========================
# PDF ROUTE
# =========================

PDF_MAP = {
    "basic": ("Basic Details","Basic Details"),
    "service": ("Service","Service"),
    "repair": ("Repair","Repair"),
    "add_work": ("Add_Work","Add Work"),
    "pending": ("Pending_Works","Pending Works"),
    "expenditure": ("Expenditure","Expenditure"),
    "tyre": ("Tyre1","Tyre"),
    "fuel": ("Fuel Cost","Fuel"),
    "revenue": ("Revnue Li","Revenue"),
    "exp_sum": ("exp_sum","Exp Summary"),
    "fuel_sum": ("Fuel Sum","Fuel Summary"),
    "prov_bal": ("prov_bal","Provision Balance")
}

# ========================= 
# PDF ROUTE 
# =========================

@app.route("/report/<key>/pdf")
def pdf(key):

    if key not in PDF_MAP:
        return "Invalid report key", 404

    sheet, title = PDF_MAP[key]

    buffer = render_pdf(sheet, title, key)

    return send_file(
        buffer,
        as_attachment=True,
        download_name=f"{key}.pdf",
        mimetype="application/pdf"
    )

# =========================
# VEHICLE GALLERY
# =========================

@app.route("/vehicles")
def vehicles():

    if require_login():
        return require_login()

    df = read_sheet("Basic Details")

    vehicle_list = (
        df["Vehicle No"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    vehicle_list.sort()

    role = session.get("role", "")
    assigned = session.get("assigned_vehicles", "")

    # If Driver or Executive → filter vehicles
    if role in ["Driver", "Executive"]:
        if assigned:
            allowed = [v.strip() for v in assigned.split(",")]
            vehicle_list = [v for v in vehicle_list if v in allowed]
        else:
            vehicle_list = []

    return render_template(
        "vehicles.html",
        vehicles=vehicle_list
    )

# =========================
# CACHE CLEAR ROUTE
# =========================

@app.route("/clear-cache")
def clear_cache():
    read_sheet.cache_clear()
    return "Google Sheet cache cleared!"

# =========================

import webbrowser
import threading

def open_browser():
    webbrowser.open("http://127.0.0.1:5000")

if __name__ == "__main__":
    threading.Timer(1.5, open_browser).start()
    app.run()
