"""
Horizon Airways - Airline Reservation System (Cross-Platform Web Version)
==========================================================================
Original project: Visual Basic .NET WinForms + MS Access (.mdb)
                  -> Windows only (WinForms and Access do not run on macOS)

This version  : Python + Flask + SQLite
                  -> Runs on Windows, macOS and Linux with the same code.
                     SQLite replaces MS Access, the browser replaces WinForms.

Run:
    Windows :  run_windows.bat        (or: py app.py)
    macOS   :  ./run_mac_linux.sh     (or: python3 app.py)
    Linux   :  ./run_mac_linux.sh     (or: python3 app.py)
Then open http://127.0.0.1:5000 in any browser.
"""

import os
import random
import sqlite3
import string
from datetime import date, datetime
from functools import wraps

from flask import Flask, flash, g, redirect, render_template_string, request, session, url_for
from werkzeug.security import check_password_hash, generate_password_hash

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATABASE = os.path.join(BASE_DIR, "airline.db")

app = Flask(__name__)
app.secret_key = "horizon-airways-dev-key-change-me"

# The owner account: logging in with this email is automatically an admin,
# so there is no separate admin signup/login page.
OWNER_EMAIL = "rishabhsingh4630@gmail.com"


# --------------------------------------------------------------------------
# Database helpers (SQLite - cross-platform replacement for MS Access .mdb)
# --------------------------------------------------------------------------
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DATABASE)
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


@app.teardown_appcontext
def close_db(_exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db(seed=True):
    db = sqlite3.connect(DATABASE)
    db.execute("PRAGMA foreign_keys = ON")
    db.executescript(
        """
        CREATE TABLE IF NOT EXISTS sectors (
            sector_id   TEXT PRIMARY KEY,
            origin      TEXT NOT NULL,
            destination TEXT NOT NULL,
            description TEXT DEFAULT ''
        );
        CREATE TABLE IF NOT EXISTS aircraft (
            aircraft_id    TEXT PRIMARY KEY,
            description    TEXT NOT NULL,
            first_seats    INTEGER NOT NULL DEFAULT 0,
            business_seats INTEGER NOT NULL DEFAULT 0,
            economy_seats  INTEGER NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS flights (
            flight_no      TEXT PRIMARY KEY,
            sector_id      TEXT NOT NULL REFERENCES sectors(sector_id),
            aircraft_id    TEXT NOT NULL REFERENCES aircraft(aircraft_id),
            dep_time       TEXT NOT NULL,
            arr_time       TEXT NOT NULL,
            weekday1       TEXT DEFAULT '',
            weekday2       TEXT DEFAULT '',
            first_fare     REAL NOT NULL DEFAULT 0,
            business_fare  REAL NOT NULL DEFAULT 0,
            economy_fare   REAL NOT NULL DEFAULT 0
        );
        CREATE TABLE IF NOT EXISTS scheduled_flights (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            flight_no  TEXT NOT NULL REFERENCES flights(flight_no),
            flight_date TEXT NOT NULL,
            UNIQUE(flight_no, flight_date)
        );
        CREATE TABLE IF NOT EXISTS reservations (
            pnr         TEXT PRIMARY KEY,
            first_name  TEXT NOT NULL,
            last_name   TEXT NOT NULL,
            age         INTEGER,
            gender      TEXT,
            travel_class TEXT NOT NULL,          -- First / Business / Economy
            flight_no   TEXT NOT NULL REFERENCES flights(flight_no),
            travel_date TEXT NOT NULL,
            fare        REAL NOT NULL DEFAULT 0,
            seat_pref   TEXT DEFAULT 'Window',   -- Window / Aisle
            meal_pref   TEXT DEFAULT 'Veg',      -- Veg / Non-veg
            ssr         TEXT DEFAULT '',         -- Special Service Request
            status      TEXT NOT NULL DEFAULT 'Confirmed',  -- Confirmed / Wait-listed / Cancelled
            created_on  TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS daily_collections (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            pnr       TEXT NOT NULL,
            tran_date TEXT NOT NULL,
            tran_type TEXT NOT NULL,             -- Payment / Refund
            amount    REAL NOT NULL
        );
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            name          TEXT NOT NULL,
            email         TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            role          TEXT NOT NULL DEFAULT 'customer',  -- customer / admin
            created_on    TEXT NOT NULL
        );
        """
    )
    # Migration for databases created before login/signup was added:
    # old reservations rows have no owner, so the new column is nullable.
    cols = [r[1] for r in db.execute("PRAGMA table_info(reservations)")]
    if "user_id" not in cols:
        db.execute("ALTER TABLE reservations ADD COLUMN user_id INTEGER REFERENCES users(id)")
    ucols = [r[1] for r in db.execute("PRAGMA table_info(users)")]
    if "role" not in ucols:
        db.execute("ALTER TABLE users ADD COLUMN role TEXT NOT NULL DEFAULT 'customer'")
    if seed and db.execute("SELECT COUNT(*) FROM sectors").fetchone()[0] == 0:
        db.executemany(
            "INSERT INTO sectors VALUES (?,?,?,?)",
            [
                ("DEL-BOM", "Delhi", "Mumbai", "Delhi to Mumbai"),
                ("DEL-BLR", "Delhi", "Bengaluru", "Delhi to Bengaluru"),
                ("BOM-GOA", "Mumbai", "Goa", "Mumbai to Goa"),
                ("LKO-DEL", "Lucknow", "Delhi", "Lucknow to Delhi"),
            ],
        )
        db.executemany(
            "INSERT INTO aircraft VALUES (?,?,?,?,?)",
            [
                ("A320", "Airbus A320", 8, 16, 120),
                ("B737", "Boeing 737", 0, 12, 150),
                ("ATR72", "ATR 72", 0, 0, 70),
            ],
        )
        db.executemany(
            "INSERT INTO flights VALUES (?,?,?,?,?,?,?,?,?,?)",
            [
                ("HZ101", "DEL-BOM", "A320", "06:00", "08:10", "Monday", "Friday", 12000, 8000, 4500),
                ("HZ202", "DEL-BLR", "B737", "09:00", "11:45", "Tuesday", "Saturday", 0, 9000, 5200),
                ("HZ303", "BOM-GOA", "ATR72", "13:00", "14:15", "Wednesday", "Sunday", 0, 0, 3200),
                ("HZ404", "LKO-DEL", "ATR72", "07:30", "08:40", "Monday", "Thursday", 0, 0, 2800),
            ],
        )
        today = date.today().isoformat()
        db.executemany(
            "INSERT OR IGNORE INTO scheduled_flights (flight_no, flight_date) VALUES (?,?)",
            [("HZ101", today), ("HZ202", today), ("HZ303", today), ("HZ404", today)],
        )
    db.commit()
    db.close()


# --------------------------------------------------------------------------
# Business logic - same rules as the original VB.NET application
# --------------------------------------------------------------------------
CLASS_SEAT_COLUMN = {"First": "first_seats", "Business": "business_seats", "Economy": "economy_seats"}
CLASS_FARE_COLUMN = {"First": "first_fare", "Business": "business_fare", "Economy": "economy_fare"}


def seats_available(db, flight_no, travel_date, travel_class):
    row = db.execute(
        """SELECT a.* FROM flights f JOIN aircraft a ON a.aircraft_id = f.aircraft_id
           WHERE f.flight_no = ?""",
        (flight_no,),
    ).fetchone()
    if row is None:
        return 0
    total = row[CLASS_SEAT_COLUMN[travel_class]]
    booked = db.execute(
        """SELECT COUNT(*) FROM reservations
           WHERE flight_no=? AND travel_date=? AND travel_class=? AND status='Confirmed'""",
        (flight_no, travel_date, travel_class),
    ).fetchone()[0]
    return total - booked


def fare_for(db, flight_no, travel_class):
    row = db.execute("SELECT * FROM flights WHERE flight_no=?", (flight_no,)).fetchone()
    return row[CLASS_FARE_COLUMN[travel_class]] if row else 0


def calculate_refund(status, amount_paid, travel_date_str):
    """Refund rule copied from CancelReservation.vb (CalculateRefundAmt):
       - Wait-listed / Overbooked  -> full refund
       - Confirmed, travel in <= 1 day -> 90% refund
       - Confirmed, otherwise      -> full refund
    """
    if status in ("Wait-listed", "Overbooked"):
        return float(amount_paid)
    try:
        travel = datetime.strptime(travel_date_str, "%Y-%m-%d").date()
    except ValueError:
        return float(amount_paid)
    days = (travel - date.today()).days
    return round(float(amount_paid) * 0.9, 2) if days <= 1 else float(amount_paid)


def new_pnr(db):
    while True:
        pnr = "".join(random.choices(string.ascii_uppercase + string.digits, k=5))
        if db.execute("SELECT 1 FROM reservations WHERE pnr=?", (pnr,)).fetchone() is None:
            return pnr


# --------------------------------------------------------------------------
# Login / Signup (Flask session; passwords are stored only as hashes,
# never in plain text - Werkzeug's generate_password_hash / check_password_hash)
# --------------------------------------------------------------------------
def current_user():
    if "user_id" not in session:
        return None
    return get_db().execute("SELECT id, name, email, role FROM users WHERE id=?", (session["user_id"],)).fetchone()


def login_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("Please login first to continue.")
            return redirect(url_for("login", next=request.path))
        return view(*args, **kwargs)
    return wrapped


def admin_required(view):
    @wraps(view)
    def wrapped(*args, **kwargs):
        if "user_id" not in session:
            flash("Please login as admin first.")
            return redirect(url_for("login", next=request.path))
        if session.get("user_role") != "admin":
            flash("Only admins can open this page.")
            return redirect(url_for("home"))
        return view(*args, **kwargs)
    return wrapped


def owns_booking(row):
    # Admins can access any booking; customers only their own.
    # Bookings made before roles existed have user_id NULL and are
    # treated as admin-visible only, so a customer cannot claim them.
    return session.get("user_role") == "admin" or (
        row["user_id"] is not None and row["user_id"] == session.get("user_id")
    )


def effective_role(user):
    return "admin" if user["email"].lower() == OWNER_EMAIL else user["role"]


def start_session(user):
    session["user_id"] = user["id"]
    session["user_name"] = user["name"]
    session["user_role"] = effective_role(user)
    if session["user_role"] == "admin" and user["role"] != "admin":
        db = get_db()
        db.execute("UPDATE users SET role='admin' WHERE id=?", (user["id"],))
        db.commit()


# --------------------------------------------------------------------------
# Templates (single-file app: layout + pages)
# --------------------------------------------------------------------------
LAYOUT = """
<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Horizon Airways - {{ title }}</title>
<link href="https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
 /* Same look as the GitHub Pages demo: warm off-white background, navy
    top bar, amber brand mark, teal buttons, soft shadows, rounded cards. */
 :root{--bg:#f6f4ef;--surface:#ffffff;--surface2:#eef2f5;--ink:#13273f;--muted:#5d6f84;--faint:#8494a7;
  --navy:#0f2e4d;--navy2:#17426b;--accent:#0e7a6d;--accent-soft:#e0f1ee;--amber:#f59e0b;
  --danger:#b42318;--danger-soft:#fdecea;--ok:#067647;--ok-soft:#e6f4ea;
  --line:#e3ded1;--line2:#dbe2ea;--shadow:0 10px 30px rgba(15,46,77,.10);--radius:16px}
 @media (prefers-color-scheme:dark){
  :root{--bg:#091423;--surface:#101f33;--surface2:#16293f;--ink:#edf3f9;--muted:#a4b6c9;--faint:#75899f;
   --navy:#0b1e33;--navy2:#14304f;--accent:#2fbfab;--accent-soft:#123a36;--amber:#f0a13c;
   --danger:#ff8a7a;--danger-soft:#3d1a16;--ok:#5fd68f;--ok-soft:#14301f;
   --line:#22344c;--line2:#22344c;--shadow:0 10px 30px rgba(0,0,0,.35)}}
 *{box-sizing:border-box}
 body{font-family:"Archivo",system-ui,-apple-system,'Segoe UI',Arial,sans-serif;margin:0;background:var(--bg);color:var(--ink);-webkit-font-smoothing:antialiased}
 a{color:var(--accent)}
 header{position:sticky;top:0;z-index:40;background:var(--navy);color:#fff;box-shadow:0 4px 18px rgba(9,20,35,.25)}
 .topbar{max-width:1180px;margin:0 auto;padding:12px 18px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
 .brand{display:flex;align-items:center;gap:11px;margin-right:auto;text-decoration:none;color:#fff}
 .brand-mark{width:38px;height:38px;border-radius:10px;background:#f59e0b;color:#12263f;display:grid;place-items:center;font-weight:900;font-size:1.25rem}
 .brand b{font-size:1.12rem;display:block;line-height:1.1}
 .brand span{font-size:.72rem;letter-spacing:.16em;text-transform:uppercase;color:#bcd2e8;font-weight:600}
 nav{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
 nav a{color:#d7e5f2;text-decoration:none;padding:8px 13px;border-radius:999px;font-weight:600;font-size:.9rem}
 nav a:hover{background:rgba(255,255,255,.10);color:#fff}
 nav a.cta{background:#f59e0b;color:#12263f}
 nav .hello{color:#c9dcee;font-size:.9rem;font-weight:600}
 main{max-width:1180px;margin:0 auto;padding:26px 18px 64px}
 .card{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);padding:20px;box-shadow:var(--shadow);margin-bottom:18px}
 .hero{background:var(--navy);color:#fff;border-radius:22px;padding:30px;margin-bottom:22px;box-shadow:var(--shadow)}
 .hero h1{font-size:clamp(1.9rem,4vw,2.9rem);line-height:1.04;margin:10px 0 12px;font-weight:900}
 .hero p{color:#c9dcee;line-height:1.6;margin:0 0 20px;max-width:46ch}
 .hero-tag{display:inline-block;background:rgba(245,158,11,.18);color:#fbbf24;border:1px solid rgba(245,158,11,.45);padding:5px 11px;border-radius:999px;font-size:.74rem;font-weight:700;letter-spacing:.14em;text-transform:uppercase}
 table{width:100%;border-collapse:collapse} th,td{padding:10px 8px;border-bottom:1px solid var(--line2);text-align:left;font-size:14px}
 th{font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;color:var(--muted)}
 input,select{font-family:inherit;font-size:1rem;color:var(--ink);background:var(--surface);padding:10px;border:1px solid var(--line2);border-radius:9px;margin:4px 8px 12px 0}
 button,.btn{font-family:inherit;border:0;border-radius:11px;padding:12px 20px;font-weight:800;font-size:.95rem;background:var(--accent);color:#fff;cursor:pointer;text-decoration:none;display:inline-flex;align-items:center;gap:8px}
 button:hover,.btn:hover{filter:brightness(1.07)}
 .btn.light,.btn.amber{background:#f59e0b;color:#12263f}
 .btn.ghost{background:transparent;border:1.5px solid var(--line2);color:var(--ink)}
 .danger{background:var(--danger)}
 label{font-size:.72rem;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);display:block}
 .flash{background:var(--ok-soft);border:1px solid var(--ok);padding:12px 14px;border-radius:11px;margin-bottom:14px}
 .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}
 .grid .card{padding:16px 16px 13px}
 .stat{font-size:1.65rem;font-weight:900;line-height:1;color:var(--navy)}
 /* Premium polish: soft entrance, hover lift, smooth buttons, focus rings */
 @keyframes rise{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
 main>*{animation:rise .45s ease both}
 main>*:nth-child(2){animation-delay:.05s} main>*:nth-child(3){animation-delay:.1s}
 .card{transition:transform .2s ease,box-shadow .2s ease}
 .card:hover{transform:translateY(-2px);box-shadow:0 16px 34px rgba(15,46,77,.16)}
 .hero{background:linear-gradient(135deg,var(--navy) 0%,var(--navy2) 60%,#1d5c54 100%);position:relative;overflow:hidden}
 .hero:after{content:"✈";position:absolute;right:18px;bottom:-24px;font-size:9rem;opacity:.08;transform:rotate(-18deg)}
 button,.btn{transition:transform .15s ease,filter .15s ease,box-shadow .15s ease;box-shadow:0 4px 14px rgba(14,122,109,.25)}
 button:hover,.btn:hover{transform:translateY(-1px)}
 button:active,.btn:active{transform:translateY(0) scale(.98)}
 input:focus,select:focus{outline:2px solid var(--accent);outline-offset:1px;border-color:var(--accent)}
 tr{transition:background .15s ease} tbody tr:hover{background:var(--surface2)}
 nav a{transition:background .15s ease,color .15s ease}
 .brand-mark{transition:transform .25s ease} .brand:hover .brand-mark{transform:rotate(-12deg) scale(1.06)}
 @media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
 @media print{header,.noprint{display:none}}
</style></head><body>
<header><div class="topbar">
 <a class="brand" href="{{ url_for('home') }}"><span class="brand-mark">✈</span><span><b>Horizon Airways</b><span>Airline Reservation System</span></span></a>
 <nav>
 <a href="{{ url_for('home') }}">Main Menu</a>
 <a href="{{ url_for('reservations') }}">Reservations</a>
 <a class="cta" href="{{ url_for('new_reservation') }}">New Reservation</a>
 <a href="{{ url_for('cancel') }}">Cancel Reservation</a>
 <a href="{{ url_for('flights') }}">Flights</a>
 <a href="{{ url_for('schedule') }}">Schedule Flights</a>
 <a href="{{ url_for('sectors') }}">Sectors</a>
 <a href="{{ url_for('aircraft') }}">Aircraft</a>
 <a href="{{ url_for('reports') }}">Reports</a>
 {% if session.get('user_id') %}
  <span class="hello">Hello, {{ session.get('user_name') }} ({{ session.get('user_role') }})</span>
  <a href="{{ url_for('logout') }}">Logout</a>
 {% else %}
  <a href="{{ url_for('login') }}">Login</a>
  <a href="{{ url_for('register') }}">Sign Up</a>
 {% endif %}
 </nav>
</div></header>
<main>
{% with messages = get_flashed_messages() %}{% if messages %}
  {% for m in messages %}<div class="flash">{{ m }}</div>{% endfor %}{% endif %}{% endwith %}
{{ body|safe }}
</main></body></html>
"""


def page(title, body, **ctx):
    return render_template_string(LAYOUT, title=title, body=render_template_string(body, **ctx))


def table_html(rows, cols):
    if not rows:
        return "<p>No records yet.</p>"
    head = "".join(f"<th>{c[1]}</th>" for c in cols)
    body = ""
    for r in rows:
        body += "<tr>" + "".join(f"<td>{'' if r[c[0]] is None else r[c[0]]}</td>" for c in cols) + "</tr>"
    return f"<table><tr>{head}</tr>{body}</table>"


# --------------------------------------------------------------------------
# Routes - Sign Up / Login / Logout
# --------------------------------------------------------------------------
@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        name = request.form["name"].strip()
        email = request.form["email"].strip().lower()
        password = request.form["password"]
        if len(password) < 6:
            flash("Password must be at least 6 characters.")
        elif get_db().execute("SELECT 1 FROM users WHERE email=?", (email,)).fetchone():
            flash("This email is already registered. Please login instead.")
            return redirect(url_for("login"))
        else:
            db = get_db()
            cur = db.execute(
                "INSERT INTO users (name, email, password_hash, role, created_on) VALUES (?,?,?,?,?)",
                (name, email, generate_password_hash(password),
                 "admin" if email == OWNER_EMAIL else "customer", date.today().isoformat()),
            )
            db.commit()
            session["user_id"] = cur.lastrowid
            session["user_name"] = name
            session["user_role"] = "admin" if email == OWNER_EMAIL else "customer"
            flash(f"Welcome, {name}! Your account has been created.")
            return redirect(url_for("home"))
    body = """
    <div class="card"><h2>Sign Up</h2><form method="post">
    <label>Full Name<input name="name" required></label>
    <label>Email<input name="email" type="email" required></label>
    <label>Password<input name="password" type="password" minlength="6" required></label>
    <button>Create Account</button></form>
    <p>Already have an account? <a href="{{ url_for('login') }}">Login</a></p></div>
    """
    return page("Sign Up", body)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].strip().lower()
        user = get_db().execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if user and check_password_hash(user["password_hash"], request.form["password"]):
            start_session(user)
            flash(f"Welcome back, {user['name']}!")
            return redirect(request.args.get("next") or url_for("home"))
        else:
            flash("Invalid email or password. Please try again.")
    body = """
    <div class="card"><h2>Login</h2><form method="post">
    <label>Email<input name="email" type="email" required></label>
    <label>Password<input name="password" type="password" required></label>
    <button>Login</button></form>
    <p>Don't have an account? <a href="{{ url_for('register') }}">Sign Up</a></p></div>
    """
    return page("Login", body)


@app.route("/logout")
def logout():
    session.clear()
    flash("You have been logged out.")
    return redirect(url_for("home"))


# --------------------------------------------------------------------------
# Routes - Main Menu / Dashboard
# --------------------------------------------------------------------------
@app.route("/")
def home():
    db = get_db()
    stats = {
        "Flights": db.execute("SELECT COUNT(*) FROM flights").fetchone()[0],
        "Sectors": db.execute("SELECT COUNT(*) FROM sectors").fetchone()[0],
        "Aircraft": db.execute("SELECT COUNT(*) FROM aircraft").fetchone()[0],
        "Reservations": db.execute("SELECT COUNT(*) FROM reservations WHERE status!='Cancelled'").fetchone()[0],
        "Confirmed": db.execute("SELECT COUNT(*) FROM reservations WHERE status='Confirmed'").fetchone()[0],
        "Wait-listed": db.execute("SELECT COUNT(*) FROM reservations WHERE status='Wait-listed'").fetchone()[0],
    }
    cards = "".join(
        f'<div class="card"><div class="stat">{v}</div><div>{k}</div></div>' for k, v in stats.items()
    )
    body = f"""
    <div class="hero"><span class="hero-tag">Horizon Airways</span>
    <h1>Book a flight in under a minute.</h1>
    <p>Bookings, reports and master tables — same features as the original VB.NET
    desktop app, now in the browser on any operating system.</p>
    <p><a class="btn light" href="{url_for('new_reservation')}">Make a Reservation</a>
    <a class="btn" style="background:#fff;color:var(--navy)" href="{url_for('flights')}">View Flights</a></p></div>
    <div class="grid">{cards}</div>
    <div class="card"><h3>Quick links</h3>
    <p><a class="btn" href="{url_for('new_reservation')}">Make a Reservation</a>
    <a class="btn ghost" href="{url_for('cancel')}">Cancel a Reservation</a>
    <a class="btn ghost" href="{url_for('reports')}">View Reports</a></p></div>
    """
    return page("Main Menu", body)


# --------------------------------------------------------------------------
# Sectors / Aircraft / Flights (Master tables)
# --------------------------------------------------------------------------
@app.route("/sectors", methods=["GET", "POST"])
def sectors():
    db = get_db()
    if request.method == "POST":
        if "user_id" not in session:
            flash("Please login first to continue.")
            return redirect(url_for("login", next=request.path))
        if session.get("user_role") != "admin":
            flash("Only admins can add or edit master data.")
            return redirect(url_for("home"))
        db.execute(
            "INSERT OR REPLACE INTO sectors VALUES (?,?,?,?)",
            (request.form["sector_id"].strip().upper(), request.form["origin"].strip(),
             request.form["destination"].strip(), request.form["description"].strip()),
        )
        db.commit()
        flash("Sector saved.")
        return redirect(url_for("sectors"))
    rows = db.execute("SELECT * FROM sectors ORDER BY sector_id").fetchall()
    body = f"""
    <div class="card"><h2>Sectors</h2>
    {table_html(rows, [('sector_id','Sector ID'),('origin','Origin'),('destination','Destination'),('description','Description')])}
    </div>
    <div class="card"><h3>Add / Edit Sector</h3><form method="post">
    <label>Sector ID (e.g. DEL-BOM)<input name="sector_id" required></label>
    <label>Origin<input name="origin" required></label>
    <label>Destination<input name="destination" required></label>
    <label>Description<input name="description"></label>
    <button>Save</button></form></div>
    """
    return page("Sectors", body)


@app.route("/aircraft", methods=["GET", "POST"])
def aircraft():
    db = get_db()
    if request.method == "POST":
        if "user_id" not in session:
            flash("Please login first to continue.")
            return redirect(url_for("login", next=request.path))
        if session.get("user_role") != "admin":
            flash("Only admins can add or edit master data.")
            return redirect(url_for("home"))
        db.execute(
            "INSERT OR REPLACE INTO aircraft VALUES (?,?,?,?,?)",
            (request.form["aircraft_id"].strip().upper(), request.form["description"].strip(),
             int(request.form["first_seats"] or 0), int(request.form["business_seats"] or 0),
             int(request.form["economy_seats"] or 0)),
        )
        db.commit()
        flash("Aircraft saved.")
        return redirect(url_for("aircraft"))
    rows = db.execute("SELECT * FROM aircraft ORDER BY aircraft_id").fetchall()
    body = f"""
    <div class="card"><h2>Aircraft</h2>
    {table_html(rows, [('aircraft_id','Aircraft Type ID'),('description','Description'),('first_seats','First Class Seats'),('business_seats','Business Class Seats'),('economy_seats','Economy Class Seats')])}
    </div>
    <div class="card"><h3>Add / Edit Aircraft</h3><form method="post">
    <label>Aircraft Type ID (e.g. A320)<input name="aircraft_id" required></label>
    <label>Description<input name="description" required></label>
    <label>First Class Seats<input name="first_seats" type="number" value="0" min="0"></label>
    <label>Business Class Seats<input name="business_seats" type="number" value="0" min="0"></label>
    <label>Economy Class Seats<input name="economy_seats" type="number" value="0" min="0"></label>
    <button>Save</button></form></div>
    """
    return page("Aircraft", body)


@app.route("/flights", methods=["GET", "POST"])
def flights():
    db = get_db()
    if request.method == "POST":
        if "user_id" not in session:
            flash("Please login first to continue.")
            return redirect(url_for("login", next=request.path))
        if session.get("user_role") != "admin":
            flash("Only admins can add or edit master data.")
            return redirect(url_for("home"))
        db.execute(
            "INSERT OR REPLACE INTO flights VALUES (?,?,?,?,?,?,?,?,?,?)",
            (request.form["flight_no"].strip().upper(), request.form["sector_id"], request.form["aircraft_id"],
             request.form["dep_time"], request.form["arr_time"], request.form["weekday1"], request.form["weekday2"],
             float(request.form["first_fare"] or 0), float(request.form["business_fare"] or 0),
             float(request.form["economy_fare"] or 0)),
        )
        db.commit()
        flash("Flight saved.")
        return redirect(url_for("flights"))
    rows = db.execute(
        "SELECT * FROM flights ORDER BY flight_no"
    ).fetchall()
    sector_opts = "".join(
        f'<option value="{r["sector_id"]}">{r["sector_id"]} ({r["origin"]} → {r["destination"]})</option>'
        for r in db.execute("SELECT * FROM sectors ORDER BY sector_id")
    )
    aircraft_opts = "".join(
        f'<option value="{r["aircraft_id"]}">{r["aircraft_id"]} — {r["description"]}</option>'
        for r in db.execute("SELECT * FROM aircraft ORDER BY aircraft_id")
    )
    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    day_opts = "".join(f'<option>{d}</option>' for d in days)
    body = f"""
    <div class="card"><h2>Flights</h2>
    {table_html(rows, [('flight_no','Flight No.'),('sector_id','Sector ID'),('aircraft_id','Aircraft Type ID'),('dep_time','Departure'),('arr_time','Arrival'),('weekday1','WeekDay 1'),('weekday2','WeekDay 2'),('first_fare','First Fare ($)'),('business_fare','Business Fare ($)'),('economy_fare','Economy Fare ($)')])}
    </div>
    <div class="card"><h3>Add / Edit Flight</h3><form method="post">
    <label>Flight Number<input name="flight_no" required placeholder="HZ101"></label>
    <label>Sector<select name="sector_id">{sector_opts}</select></label>
    <label>Aircraft<select name="aircraft_id">{aircraft_opts}</select></label>
    <label>Departure Time<input name="dep_time" type="time" required></label>
    <label>Arrival Time<input name="arr_time" type="time" required></label>
    <label>WeekDay 1<select name="weekday1">{day_opts}</select></label>
    <label>WeekDay 2<select name="weekday2">{day_opts}</select></label>
    <label>First Class Fare ($)<input name="first_fare" type="number" step="0.01" value="0"></label>
    <label>Business Class Fare ($)<input name="business_fare" type="number" step="0.01" value="0"></label>
    <label>Economy Class Fare ($)<input name="economy_fare" type="number" step="0.01" value="0"></label>
    <button>Save</button></form></div>
    """
    return page("Flights", body)


@app.route("/schedule", methods=["GET", "POST"])
def schedule():
    db = get_db()
    if request.method == "POST":
        if "user_id" not in session:
            flash("Please login first to continue.")
            return redirect(url_for("login", next=request.path))
        if session.get("user_role") != "admin":
            flash("Only admins can add or edit master data.")
            return redirect(url_for("home"))
        db.execute(
            "INSERT OR IGNORE INTO scheduled_flights (flight_no, flight_date) VALUES (?,?)",
            (request.form["flight_no"], request.form["flight_date"]),
        )
        db.commit()
        flash("Flight scheduled.")
        return redirect(url_for("schedule"))
    rows = db.execute(
        """SELECT sf.id, sf.flight_no, sf.flight_date, f.dep_time, f.arr_time, s.origin, s.destination
           FROM scheduled_flights sf JOIN flights f ON f.flight_no=sf.flight_no
           JOIN sectors s ON s.sector_id=f.sector_id ORDER BY sf.flight_date, sf.flight_no"""
    ).fetchall()
    flight_opts = "".join(
        f'<option value="{r["flight_no"]}">{r["flight_no"]}</option>'
        for r in db.execute("SELECT flight_no FROM flights ORDER BY flight_no")
    )
    body = f"""
    <div class="card"><h2>Scheduled Flights</h2>
    {table_html(rows, [('flight_no','Flight No.'),('flight_date','Flight Date'),('origin','From'),('destination','To'),('dep_time','Departure'),('arr_time','Arrival')])}
    </div>
    <div class="card"><h3>Schedule a Flight</h3><form method="post">
    <label>Flight<select name="flight_no">{flight_opts}</select></label>
    <label>Flight Date<input name="flight_date" type="date" required></label>
    <button>Schedule</button></form></div>
    """
    return page("Schedule Flights", body)


# --------------------------------------------------------------------------
# Reservations
# --------------------------------------------------------------------------
@app.route("/reservations")
@login_required
def reservations():
    db = get_db()
    if session.get("user_role") == "admin" and request.args.get("mine") != "1":
        rows = db.execute(
            """SELECT pnr, (first_name||' '||last_name) AS passenger, age, gender, travel_class,
                      flight_no, travel_date, fare, status FROM reservations ORDER BY created_on DESC"""
        ).fetchall()
    else:
        rows = db.execute(
            """SELECT pnr, (first_name||' '||last_name) AS passenger, age, gender, travel_class,
                      flight_no, travel_date, fare, status FROM reservations
               WHERE user_id=? ORDER BY created_on DESC""",
            (session["user_id"],),
        ).fetchall()
    extra = ""
    if rows:
        links = " ".join(
            f'<div><a href="{url_for("ticket", pnr=r["pnr"])}">Print ticket — {r["pnr"]}</a></div>' for r in rows
        )
        extra = f'<div class="card"><h3>Print Tickets</h3>{links}</div>'
    body = f"""
    <div class="card"><h2>Reservations</h2>
    {table_html(rows, [('pnr','PNR No.'),('passenger','Passenger Name'),('age','Age'),('gender','Gender'),('travel_class','Class'),('flight_no','Flight No.'),('travel_date','Travel Date'),('fare','Fare Amount ($)'),('status','Status')])}
    <p><a class="btn" href="{url_for('new_reservation')}">New Reservation</a>
    <a class="btn light" href="{url_for('reservations', mine=1)}">My Bookings</a>
    <a class="btn light" href="{url_for('reservations')}">All Bookings</a></p></div>{extra}
    """
    return page("Reservations", body)


@app.route("/reservations/new", methods=["GET", "POST"])
@login_required
def new_reservation():
    db = get_db()
    if request.method == "POST":
        f = request.form
        travel_class = f["travel_class"]
        fare = fare_for(db, f["flight_no"], travel_class)
        available = seats_available(db, f["flight_no"], f["travel_date"], travel_class)
        status = "Confirmed" if available > 0 else "Wait-listed"
        pnr = new_pnr(db)
        db.execute(
            "INSERT INTO reservations VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
            (pnr, f["first_name"].strip(), f["last_name"].strip(), int(f["age"] or 0), f["gender"],
             travel_class, f["flight_no"], f["travel_date"], fare, f["seat_pref"], f["meal_pref"],
             f["ssr"].strip(), status, date.today().isoformat(), session.get("user_id")),
        )
        db.execute(
            "INSERT INTO daily_collections (pnr, tran_date, tran_type, amount) VALUES (?,?,?,?)",
            (pnr, date.today().isoformat(), "Payment", fare),
        )
        db.commit()
        flash(f"Reservation created. PNR: {pnr} — Status: {status} — Fare: $ {fare:.2f}")
        return redirect(url_for("ticket", pnr=pnr))
    flight_opts = "".join(
        f'<option value="{r["flight_no"]}">{r["flight_no"]} — {r["origin"]} → {r["destination"]} ({r["dep_time"]})</option>'
        for r in db.execute(
            """SELECT f.flight_no, f.dep_time, s.origin, s.destination FROM flights f
               JOIN sectors s ON s.sector_id=f.sector_id ORDER BY f.flight_no"""
        )
    )
    body = f"""
    <div class="card"><h2>New Reservation</h2><form method="post">
    <h3>Passenger Details</h3>
    <label>First Name<input name="first_name" required></label>
    <label>Last Name<input name="last_name" required></label>
    <label>Age<input name="age" type="number" min="0" max="120" required></label>
    <label>Gender<select name="gender"><option>Male</option><option>Female</option><option>Other</option></select></label>
    <h3>Flight Details</h3>
    <label>Flight<select name="flight_no">{flight_opts}</select></label>
    <label>Travel Date<input name="travel_date" type="date" required></label>
    <label>Class<select name="travel_class"><option>Economy</option><option>Business</option><option>First</option></select></label>
    <label>Seat Preference<select name="seat_pref"><option>Window</option><option>Aisle</option></select></label>
    <label>Meal Preference<select name="meal_pref"><option>Veg</option><option>Non-veg</option></select></label>
    <label>Special Service Request<input name="ssr" placeholder="e.g. Wheelchair"></label>
    <p><button>Save Reservation &amp; Print Ticket</button></p>
    <p style="font-size:13px;color:#555">Fare is taken automatically from the flight's class fare.
    If no seat is available in the chosen class, the status becomes <b>Wait-listed</b>
    (same seat-availability check as the original app).</p>
    </form></div>
    """
    return page("New Reservation", body)


@app.route("/ticket/<pnr>")
@login_required
def ticket(pnr):
    db = get_db()
    r = db.execute(
        """SELECT res.*, s.origin, s.destination, f.dep_time, f.arr_time
           FROM reservations res JOIN flights f ON f.flight_no=res.flight_no
           JOIN sectors s ON s.sector_id=f.sector_id WHERE res.pnr=?""",
        (pnr.upper(),),
    ).fetchone()
    if r is None:
        flash("PNR not found.")
        return redirect(url_for("reservations"))
    if not owns_booking(r):
        flash("You can only view your own tickets.")
        return redirect(url_for("reservations"))
    body = f"""
    <div class="card"><h2 style="text-align:center">HORIZON AIRWAYS — E-Ticket</h2>
    <table>
     <tr><th>PNR No.</th><td>{r['pnr']}</td><th>Status</th><td>{r['status']}</td></tr>
     <tr><th>Passenger</th><td>{r['first_name']} {r['last_name']}</td><th>Age / Gender</th><td>{r['age']} / {r['gender']}</td></tr>
     <tr><th>Flight</th><td>{r['flight_no']} — {r['origin']} → {r['destination']}</td><th>Class</th><td>{r['travel_class']}</td></tr>
     <tr><th>Travel Date</th><td>{r['travel_date']}</td><th>Time</th><td>{r['dep_time']} → {r['arr_time']}</td></tr>
     <tr><th>Seat Preference</th><td>{r['seat_pref']}</td><th>Meal</th><td>{r['meal_pref']}</td></tr>
     <tr><th>Special Request</th><td>{r['ssr'] or '-'}</td><th>Fare Amount</th><td>$ {r['fare']:.2f}</td></tr>
    </table>
    <p class="noprint"><button onclick="window.print()">Print Ticket</button>
    <a class="btn light" href="{url_for('reservations')}">Back to Reservations</a></p></div>
    """
    return page("Print Ticket", body)


# --------------------------------------------------------------------------
# Cancellation + Refund receipt
# --------------------------------------------------------------------------
@app.route("/cancel", methods=["GET", "POST"])
@login_required
def cancel():
    db = get_db()
    detail = None
    refund = None
    if request.method == "POST":
        pnr = request.form["pnr"].strip().upper()
        r = db.execute("SELECT * FROM reservations WHERE pnr=?", (pnr,)).fetchone()
        if r is not None and not owns_booking(r):
            flash("You can only cancel your own reservations.")
        elif r is None:
            flash("PNR Number not found. Please re-enter.")
        elif r["status"] == "Cancelled":
            flash("The ticket is already cancelled.")
        elif r["travel_date"] < date.today().isoformat():
            flash("Cannot cancel: the travel date is before the current date.")
        elif request.form.get("action") == "cancel":
            refund = calculate_refund(r["status"], r["fare"], r["travel_date"])
            db.execute("UPDATE reservations SET status='Cancelled' WHERE pnr=?", (pnr,))
            db.execute(
                "INSERT INTO daily_collections (pnr, tran_date, tran_type, amount) VALUES (?,?,?,?)",
                (pnr, date.today().isoformat(), "Refund", refund),
            )
            # Original rule: cancelling a Confirmed Economy seat promotes the
            # earliest Wait-listed passenger on that flight/date to Confirmed.
            if r["travel_class"] == "Economy" and r["status"] == "Confirmed":
                nxt = db.execute(
                    """SELECT pnr FROM reservations WHERE flight_no=? AND travel_date=?
                       AND status='Wait-listed' ORDER BY created_on, pnr LIMIT 1""",
                    (r["flight_no"], r["travel_date"]),
                ).fetchone()
                if nxt:
                    db.execute("UPDATE reservations SET status='Confirmed' WHERE pnr=?", (nxt["pnr"],))
            db.commit()
            return redirect(url_for("refund_receipt", pnr=pnr, refund=refund))
        else:
            detail = r
            refund = calculate_refund(detail["status"], detail["fare"], detail["travel_date"])
    detail_html = ""
    if detail:
        detail_html = f"""
        <div class="card"><h3>Passenger Details</h3>
        <table>
         <tr><th>PNR</th><td>{detail['pnr']}</td><th>Passenger</th><td>{detail['first_name']} {detail['last_name']}</td></tr>
         <tr><th>Flight</th><td>{detail['flight_no']}</td><th>Travel Date</th><td>{detail['travel_date']}</td></tr>
         <tr><th>Class</th><td>{detail['travel_class']}</td><th>Status</th><td>{detail['status']}</td></tr>
         <tr><th>Amount Paid</th><td>$ {detail['fare']:.2f}</td><th>Refund Amount</th><td>$ {refund:.2f}</td></tr>
        </table>
        <form method="post"><input type="hidden" name="pnr" value="{detail['pnr']}">
        <input type="hidden" name="action" value="cancel">
        <p><button class="danger">Cancel Reservation</button></p></form></div>
        """
    body = f"""
    <div class="card"><h2>Cancel Reservation</h2>
    <form method="post"><label>Enter PNR No.<input name="pnr" maxlength="5" required></label>
    <button>Get Details</button></form></div>{detail_html}
    """
    return page("Cancel Reservation", body)


@app.route("/refund/<pnr>")
@login_required
def refund_receipt(pnr):
    refund = request.args.get("refund", "0")
    db = get_db()
    r = db.execute("SELECT * FROM reservations WHERE pnr=?", (pnr.upper(),)).fetchone()
    if r is None:
        flash("PNR not found.")
        return redirect(url_for("home"))
    if not owns_booking(r):
        flash("You can only view your own refund receipts.")
        return redirect(url_for("reservations"))
    body = f"""
    <div class="card"><h2 style="text-align:center">HORIZON AIRWAYS — Refund Receipt</h2>
    <table>
     <tr><th>PNR No.</th><td>{r['pnr']}</td><th>Passenger</th><td>{r['first_name']} {r['last_name']}</td></tr>
     <tr><th>Flight</th><td>{r['flight_no']}</td><th>Travel Date</th><td>{r['travel_date']}</td></tr>
     <tr><th>Amount Paid</th><td>$ {r['fare']:.2f}</td><th>Refund Amount</th><td>$ {float(refund):.2f}</td></tr>
     <tr><th>Status</th><td colspan="3">{r['status']}</td></tr>
    </table>
    <p class="noprint"><button onclick="window.print()">Print Refund Receipt</button></p></div>
    """
    return page("Print Refund Receipt", body)


# --------------------------------------------------------------------------
# Reports (from the original Main Menu -> Reports)
# --------------------------------------------------------------------------
@app.route("/reports")
@admin_required
def reports():
    db = get_db()
    confirmed = db.execute(
        "SELECT * FROM reservations WHERE status='Confirmed' ORDER BY travel_date, flight_no"
    ).fetchall()
    waiting = db.execute(
        "SELECT * FROM reservations WHERE status IN ('Wait-listed','Overbooked') ORDER BY travel_date, flight_no"
    ).fetchall()
    collections = db.execute(
        "SELECT * FROM daily_collections ORDER BY tran_date DESC, id DESC"
    ).fetchall()
    cols = [("pnr", "PNR"), ("first_name", "First Name"), ("last_name", "Last Name"),
            ("travel_class", "Class"), ("flight_no", "Flight"), ("travel_date", "Travel Date"),
            ("fare", "Fare ($)"), ("status", "Status")]
    body = f"""
    <div class="card"><h2>Confirmed Passenger List</h2>{table_html(confirmed, cols)}</div>
    <div class="card"><h2>Overbooking / Waiting List</h2>{table_html(waiting, cols)}</div>
    <div class="card"><h2>Daily Collections Report</h2>
    {table_html(collections, [('tran_date','Date'),('pnr','PNR'),('tran_type','Type'),('amount','Amount ($)')])}</div>
    """
    return page("Reports", body)


if __name__ == "__main__":
    init_db()
    # 0.0.0.0 lets other devices on the same network open it too.
    app.run(host="0.0.0.0", port=5000, debug=True)
else:
    init_db()
