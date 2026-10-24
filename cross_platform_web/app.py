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

    # All-India expansion (user-supplied airport/route list, 2026-10-24):
    # INSERT OR IGNORE so existing databases also gain the new routes
    # without touching the original 4 flights, and demo fares are labelled
    # as demo in the UI (these are NOT live airline fares/timings).
    db.executemany(
        "INSERT OR IGNORE INTO sectors VALUES (?,?,?,?)",
        [
                ("DEL-BOM", "Delhi", "Mumbai", "Delhi to Mumbai"),
                ("DEL-LKO", "Delhi", "Lucknow", "Delhi to Lucknow"),
                ("DEL-BLR", "Delhi", "Bengaluru", "Delhi to Bengaluru"),
                ("DEL-HYD", "Delhi", "Hyderabad", "Delhi to Hyderabad"),
                ("DEL-MAA", "Delhi", "Chennai", "Delhi to Chennai"),
                ("DEL-CCU", "Delhi", "Kolkata", "Delhi to Kolkata"),
                ("DEL-JAI", "Delhi", "Jaipur", "Delhi to Jaipur"),
                ("DEL-PAT", "Delhi", "Patna", "Delhi to Patna"),
                ("BOM-DEL", "Mumbai", "Delhi", "Mumbai to Delhi"),
                ("BOM-LKO", "Mumbai", "Lucknow", "Mumbai to Lucknow"),
                ("BOM-AMD", "Mumbai", "Ahmedabad", "Mumbai to Ahmedabad"),
                ("BOM-GOA", "Mumbai", "Goa", "Mumbai to Goa"),
                ("BOM-BLR", "Mumbai", "Bengaluru", "Mumbai to Bengaluru"),
                ("BOM-COK", "Mumbai", "Kochi", "Mumbai to Kochi"),
                ("BOM-CCU", "Mumbai", "Kolkata", "Mumbai to Kolkata"),
                ("LKO-DEL", "Lucknow", "Delhi", "Lucknow to Delhi"),
                ("LKO-BOM", "Lucknow", "Mumbai", "Lucknow to Mumbai"),
                ("LKO-BLR", "Lucknow", "Bengaluru", "Lucknow to Bengaluru"),
                ("LKO-HYD", "Lucknow", "Hyderabad", "Lucknow to Hyderabad"),
                ("LKO-CCU", "Lucknow", "Kolkata", "Lucknow to Kolkata"),
                ("LKO-AMD", "Lucknow", "Ahmedabad", "Lucknow to Ahmedabad"),
                ("BLR-DEL", "Bengaluru", "Delhi", "Bengaluru to Delhi"),
                ("BLR-BOM", "Bengaluru", "Mumbai", "Bengaluru to Mumbai"),
                ("BLR-MAA", "Bengaluru", "Chennai", "Bengaluru to Chennai"),
                ("BLR-HYD", "Bengaluru", "Hyderabad", "Bengaluru to Hyderabad"),
                ("BLR-PNQ", "Bengaluru", "Pune", "Bengaluru to Pune"),
                ("BLR-COK", "Bengaluru", "Kochi", "Bengaluru to Kochi"),
                ("BLR-CCU", "Bengaluru", "Kolkata", "Bengaluru to Kolkata"),
                ("MAA-DEL", "Chennai", "Delhi", "Chennai to Delhi"),
                ("MAA-BOM", "Chennai", "Mumbai", "Chennai to Mumbai"),
                ("MAA-BLR", "Chennai", "Bengaluru", "Chennai to Bengaluru"),
                ("MAA-HYD", "Chennai", "Hyderabad", "Chennai to Hyderabad"),
                ("MAA-CCU", "Chennai", "Kolkata", "Chennai to Kolkata"),
                ("MAA-COK", "Chennai", "Kochi", "Chennai to Kochi"),
                ("CCU-DEL", "Kolkata", "Delhi", "Kolkata to Delhi"),
                ("CCU-BOM", "Kolkata", "Mumbai", "Kolkata to Mumbai"),
                ("CCU-LKO", "Kolkata", "Lucknow", "Kolkata to Lucknow"),
                ("CCU-BLR", "Kolkata", "Bengaluru", "Kolkata to Bengaluru"),
                ("CCU-MAA", "Kolkata", "Chennai", "Kolkata to Chennai"),
                ("CCU-GAU", "Kolkata", "Guwahati", "Kolkata to Guwahati"),
                ("HYD-DEL", "Hyderabad", "Delhi", "Hyderabad to Delhi"),
                ("JAI-DEL", "Jaipur", "Delhi", "Jaipur to Delhi"),
                ("PAT-DEL", "Patna", "Delhi", "Patna to Delhi"),
                ("AMD-BOM", "Ahmedabad", "Mumbai", "Ahmedabad to Mumbai"),
                ("GOA-BOM", "Goa", "Mumbai", "Goa to Mumbai"),
                ("COK-BOM", "Kochi", "Mumbai", "Kochi to Mumbai"),
                ("BLR-LKO", "Bengaluru", "Lucknow", "Bengaluru to Lucknow"),
                ("HYD-LKO", "Hyderabad", "Lucknow", "Hyderabad to Lucknow"),
                ("AMD-LKO", "Ahmedabad", "Lucknow", "Ahmedabad to Lucknow"),
                ("HYD-BLR", "Hyderabad", "Bengaluru", "Hyderabad to Bengaluru"),
                ("PNQ-BLR", "Pune", "Bengaluru", "Pune to Bengaluru"),
                ("COK-BLR", "Kochi", "Bengaluru", "Kochi to Bengaluru"),
                ("BOM-MAA", "Mumbai", "Chennai", "Mumbai to Chennai"),
                ("HYD-MAA", "Hyderabad", "Chennai", "Hyderabad to Chennai"),
                ("COK-MAA", "Kochi", "Chennai", "Kochi to Chennai"),
                ("GAU-CCU", "Guwahati", "Kolkata", "Guwahati to Kolkata"),
                ("DEL-VNS", "Delhi", "Varanasi", "Delhi to Varanasi"),
                ("VNS-DEL", "Varanasi", "Delhi", "Varanasi to Delhi"),
                ("DEL-SXR", "Delhi", "Srinagar", "Delhi to Srinagar"),
                ("SXR-DEL", "Srinagar", "Delhi", "Srinagar to Delhi"),
                ("DEL-DED", "Delhi", "Dehradun", "Delhi to Dehradun"),
                ("DED-DEL", "Dehradun", "Delhi", "Dehradun to Delhi"),
                ("DEL-IXC", "Delhi", "Chandigarh", "Delhi to Chandigarh"),
                ("IXC-DEL", "Chandigarh", "Delhi", "Chandigarh to Delhi"),
                ("DEL-BBI", "Delhi", "Bhubaneswar", "Delhi to Bhubaneswar"),
                ("BBI-DEL", "Bhubaneswar", "Delhi", "Bhubaneswar to Delhi"),
                ("DEL-IDR", "Delhi", "Indore", "Delhi to Indore"),
                ("IDR-DEL", "Indore", "Delhi", "Indore to Delhi"),
                ("DEL-NAG", "Delhi", "Nagpur", "Delhi to Nagpur"),
                ("NAG-DEL", "Nagpur", "Delhi", "Nagpur to Delhi"),
                ("DEL-IXR", "Delhi", "Ranchi", "Delhi to Ranchi"),
                ("IXR-DEL", "Ranchi", "Delhi", "Ranchi to Delhi"),
                ("DEL-RPR", "Delhi", "Raipur", "Delhi to Raipur"),
                ("RPR-DEL", "Raipur", "Delhi", "Raipur to Delhi"),
                ("DEL-ATQ", "Delhi", "Amritsar", "Delhi to Amritsar"),
                ("ATQ-DEL", "Amritsar", "Delhi", "Amritsar to Delhi"),
                ("DEL-CJB", "Delhi", "Coimbatore", "Delhi to Coimbatore"),
                ("CJB-DEL", "Coimbatore", "Delhi", "Coimbatore to Delhi"),
                ("DEL-VTZ", "Delhi", "Visakhapatnam", "Delhi to Visakhapatnam"),
                ("VTZ-DEL", "Visakhapatnam", "Delhi", "Visakhapatnam to Delhi"),
                ("DEL-TIR", "Delhi", "Tirupati", "Delhi to Tirupati"),
                ("TIR-DEL", "Tirupati", "Delhi", "Tirupati to Delhi"),
                ("DEL-IXA", "Delhi", "Agartala", "Delhi to Agartala"),
                ("IXA-DEL", "Agartala", "Delhi", "Agartala to Delhi"),
                ("DEL-IXZ", "Delhi", "Port Blair", "Delhi to Port Blair"),
                ("IXZ-DEL", "Port Blair", "Delhi", "Port Blair to Delhi"),
        ],
    )
    db.executemany(
        "INSERT OR IGNORE INTO flights VALUES (?,?,?,?,?,?,?,?,?,?)",
        [
                ("HZ501", "DEL-BOM", "A320", "05:00", "06:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ502", "DEL-LKO", "B737", "06:35", "08:45", "Tuesday", "Friday", 0, 6100, 3400),
                ("HZ503", "DEL-BLR", "ATR72", "08:10", "10:40", "Wednesday", "Saturday", 0, 0, 4200),
                ("HZ504", "DEL-HYD", "A320", "09:45", "12:35", "Sunday", "Wednesday", 13500, 9400, 5200),
                ("HZ505", "DEL-MAA", "B737", "11:20", "14:35", "Monday", "Thursday", 0, 11500, 6400),
                ("HZ506", "DEL-CCU", "ATR72", "12:55", "16:30", "Tuesday", "Friday", 0, 0, 7800),
                ("HZ507", "DEL-JAI", "A320", "14:30", "16:20", "Wednesday", "Saturday", 7300, 5000, 2800),
                ("HZ508", "DEL-PAT", "B737", "16:05", "18:15", "Sunday", "Wednesday", 0, 6100, 3400),
                ("HZ509", "BOM-DEL", "ATR72", "17:40", "20:10", "Monday", "Thursday", 0, 0, 4200),
                ("HZ510", "BOM-LKO", "A320", "19:15", "22:05", "Tuesday", "Friday", 13500, 9400, 5200),
                ("HZ511", "BOM-AMD", "B737", "05:50", "09:05", "Wednesday", "Saturday", 0, 11500, 6400),
                ("HZ512", "BOM-GOA", "ATR72", "07:25", "11:00", "Sunday", "Wednesday", 0, 0, 7800),
                ("HZ513", "BOM-BLR", "A320", "09:00", "10:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ514", "BOM-COK", "B737", "10:35", "12:45", "Tuesday", "Friday", 0, 6100, 3400),
                ("HZ515", "BOM-CCU", "ATR72", "12:10", "14:40", "Wednesday", "Saturday", 0, 0, 4200),
                ("HZ516", "LKO-DEL", "A320", "13:45", "16:35", "Sunday", "Wednesday", 13500, 9400, 5200),
                ("HZ517", "LKO-BOM", "B737", "15:20", "18:35", "Monday", "Thursday", 0, 11500, 6400),
                ("HZ518", "LKO-BLR", "ATR72", "16:55", "20:30", "Tuesday", "Friday", 0, 0, 7800),
                ("HZ519", "LKO-HYD", "A320", "18:30", "20:20", "Wednesday", "Saturday", 7300, 5000, 2800),
                ("HZ520", "LKO-CCU", "B737", "05:05", "07:15", "Sunday", "Wednesday", 0, 6100, 3400),
                ("HZ521", "LKO-AMD", "ATR72", "06:40", "09:10", "Monday", "Thursday", 0, 0, 4200),
                ("HZ522", "BLR-DEL", "A320", "08:15", "11:05", "Tuesday", "Friday", 13500, 9400, 5200),
                ("HZ523", "BLR-BOM", "B737", "09:50", "13:05", "Wednesday", "Saturday", 0, 11500, 6400),
                ("HZ524", "BLR-MAA", "ATR72", "11:25", "15:00", "Sunday", "Wednesday", 0, 0, 7800),
                ("HZ525", "BLR-HYD", "A320", "13:00", "14:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ526", "BLR-PNQ", "B737", "14:35", "16:45", "Tuesday", "Friday", 0, 6100, 3400),
                ("HZ527", "BLR-COK", "ATR72", "16:10", "18:40", "Wednesday", "Saturday", 0, 0, 4200),
                ("HZ528", "BLR-CCU", "A320", "17:45", "20:35", "Sunday", "Wednesday", 13500, 9400, 5200),
                ("HZ529", "MAA-DEL", "B737", "19:20", "22:35", "Monday", "Thursday", 0, 11500, 6400),
                ("HZ530", "MAA-BOM", "ATR72", "05:55", "09:30", "Tuesday", "Friday", 0, 0, 7800),
                ("HZ531", "MAA-BLR", "A320", "07:30", "09:20", "Wednesday", "Saturday", 7300, 5000, 2800),
                ("HZ532", "MAA-HYD", "B737", "09:05", "11:15", "Sunday", "Wednesday", 0, 6100, 3400),
                ("HZ533", "MAA-CCU", "ATR72", "10:40", "13:10", "Monday", "Thursday", 0, 0, 4200),
                ("HZ534", "MAA-COK", "A320", "12:15", "15:05", "Tuesday", "Friday", 13500, 9400, 5200),
                ("HZ535", "CCU-DEL", "B737", "13:50", "17:05", "Wednesday", "Saturday", 0, 11500, 6400),
                ("HZ536", "CCU-BOM", "ATR72", "15:25", "19:00", "Sunday", "Wednesday", 0, 0, 7800),
                ("HZ537", "CCU-LKO", "A320", "17:00", "18:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ538", "CCU-BLR", "B737", "18:35", "20:45", "Tuesday", "Friday", 0, 6100, 3400),
                ("HZ539", "CCU-MAA", "ATR72", "05:10", "07:40", "Wednesday", "Saturday", 0, 0, 4200),
                ("HZ540", "CCU-GAU", "A320", "06:45", "09:35", "Sunday", "Wednesday", 13500, 9400, 5200),
                ("HZ541", "HYD-DEL", "B737", "08:20", "11:35", "Monday", "Thursday", 0, 11500, 6400),
                ("HZ542", "JAI-DEL", "ATR72", "09:55", "13:30", "Tuesday", "Friday", 0, 0, 7800),
                ("HZ543", "PAT-DEL", "A320", "11:30", "13:20", "Wednesday", "Saturday", 7300, 5000, 2800),
                ("HZ544", "AMD-BOM", "B737", "13:05", "15:15", "Sunday", "Wednesday", 0, 6100, 3400),
                ("HZ545", "GOA-BOM", "ATR72", "14:40", "17:10", "Monday", "Thursday", 0, 0, 4200),
                ("HZ546", "COK-BOM", "A320", "16:15", "19:05", "Tuesday", "Friday", 13500, 9400, 5200),
                ("HZ547", "BLR-LKO", "B737", "17:50", "21:05", "Wednesday", "Saturday", 0, 11500, 6400),
                ("HZ548", "HYD-LKO", "ATR72", "19:25", "23:00", "Sunday", "Wednesday", 0, 0, 7800),
                ("HZ549", "AMD-LKO", "A320", "06:00", "07:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ550", "HYD-BLR", "B737", "07:35", "09:45", "Tuesday", "Friday", 0, 6100, 3400),
                ("HZ551", "PNQ-BLR", "ATR72", "09:10", "11:40", "Wednesday", "Saturday", 0, 0, 4200),
                ("HZ552", "COK-BLR", "A320", "10:45", "13:35", "Sunday", "Wednesday", 13500, 9400, 5200),
                ("HZ553", "BOM-MAA", "B737", "12:20", "15:35", "Monday", "Thursday", 0, 11500, 6400),
                ("HZ554", "HYD-MAA", "ATR72", "13:55", "17:30", "Tuesday", "Friday", 0, 0, 7800),
                ("HZ555", "COK-MAA", "A320", "15:30", "17:20", "Wednesday", "Saturday", 7300, 5000, 2800),
                ("HZ556", "GAU-CCU", "B737", "17:05", "19:15", "Sunday", "Wednesday", 0, 6100, 3400),
                ("HZ557", "DEL-VNS", "ATR72", "18:40", "21:10", "Monday", "Thursday", 0, 0, 4200),
                ("HZ558", "VNS-DEL", "A320", "05:15", "08:05", "Tuesday", "Friday", 13500, 9400, 5200),
                ("HZ559", "DEL-SXR", "B737", "06:50", "10:05", "Wednesday", "Saturday", 0, 11500, 6400),
                ("HZ560", "SXR-DEL", "ATR72", "08:25", "12:00", "Sunday", "Wednesday", 0, 0, 7800),
                ("HZ561", "DEL-DED", "A320", "10:00", "11:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ562", "DED-DEL", "B737", "11:35", "13:45", "Tuesday", "Friday", 0, 6100, 3400),
                ("HZ563", "DEL-IXC", "ATR72", "13:10", "15:40", "Wednesday", "Saturday", 0, 0, 4200),
                ("HZ564", "IXC-DEL", "A320", "14:45", "17:35", "Sunday", "Wednesday", 13500, 9400, 5200),
                ("HZ565", "DEL-BBI", "B737", "16:20", "19:35", "Monday", "Thursday", 0, 11500, 6400),
                ("HZ566", "BBI-DEL", "ATR72", "17:55", "21:30", "Tuesday", "Friday", 0, 0, 7800),
                ("HZ567", "DEL-IDR", "A320", "19:30", "21:20", "Wednesday", "Saturday", 7300, 5000, 2800),
                ("HZ568", "IDR-DEL", "B737", "06:05", "08:15", "Sunday", "Wednesday", 0, 6100, 3400),
                ("HZ569", "DEL-NAG", "ATR72", "07:40", "10:10", "Monday", "Thursday", 0, 0, 4200),
                ("HZ570", "NAG-DEL", "A320", "09:15", "12:05", "Tuesday", "Friday", 13500, 9400, 5200),
                ("HZ571", "DEL-IXR", "B737", "10:50", "14:05", "Wednesday", "Saturday", 0, 11500, 6400),
                ("HZ572", "IXR-DEL", "ATR72", "12:25", "16:00", "Sunday", "Wednesday", 0, 0, 7800),
                ("HZ573", "DEL-RPR", "A320", "14:00", "15:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ574", "RPR-DEL", "B737", "15:35", "17:45", "Tuesday", "Friday", 0, 6100, 3400),
                ("HZ575", "DEL-ATQ", "ATR72", "17:10", "19:40", "Wednesday", "Saturday", 0, 0, 4200),
                ("HZ576", "ATQ-DEL", "A320", "18:45", "21:35", "Sunday", "Wednesday", 13500, 9400, 5200),
                ("HZ577", "DEL-CJB", "B737", "05:20", "08:35", "Monday", "Thursday", 0, 11500, 6400),
                ("HZ578", "CJB-DEL", "ATR72", "06:55", "10:30", "Tuesday", "Friday", 0, 0, 7800),
                ("HZ579", "DEL-VTZ", "A320", "08:30", "10:20", "Wednesday", "Saturday", 7300, 5000, 2800),
                ("HZ580", "VTZ-DEL", "B737", "10:05", "12:15", "Sunday", "Wednesday", 0, 6100, 3400),
                ("HZ581", "DEL-TIR", "ATR72", "11:40", "14:10", "Monday", "Thursday", 0, 0, 4200),
                ("HZ582", "TIR-DEL", "A320", "13:15", "16:05", "Tuesday", "Friday", 13500, 9400, 5200),
                ("HZ583", "DEL-IXA", "B737", "14:50", "18:05", "Wednesday", "Saturday", 0, 11500, 6400),
                ("HZ584", "IXA-DEL", "ATR72", "16:25", "20:00", "Sunday", "Wednesday", 0, 0, 7800),
                ("HZ585", "DEL-IXZ", "A320", "18:00", "19:50", "Monday", "Thursday", 7300, 5000, 2800),
                ("HZ586", "IXZ-DEL", "B737", "19:35", "21:45", "Tuesday", "Friday", 0, 6100, 3400),
        ],
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
 :root{--bg:#f4f8fd;--surface:#ffffff;--surface2:#eaf5ff;--ink:#13273f;--muted:#5d6f84;--faint:#8494a7;
  --navy:#063b78;--navy2:#0b5fa5;--accent:#0878d1;--accent-soft:#eaf5ff;--amber:#f59e0b;
  --danger:#b42318;--danger-soft:#fdecea;--ok:#067647;--ok-soft:#e6f4ea;
  --line:#dbe7f3;--line2:#dbe2ea;--shadow:0 10px 30px rgba(6,59,120,.10);--radius:16px}
 @media (prefers-color-scheme:dark){
  :root{--bg:#091423;--surface:#101f33;--surface2:#16293f;--ink:#edf3f9;--muted:#a4b6c9;--faint:#75899f;
   --navy:#041a33;--navy2:#0a3a6b;--accent:#4aa8ec;--accent-soft:#0d2a4a;--amber:#f0a13c;
   --danger:#ff8a7a;--danger-soft:#3d1a16;--ok:#5fd68f;--ok-soft:#14301f;
   --line:#2f4566;--line2:#2f4566;--shadow:0 10px 30px rgba(0,0,0,.35)}}
 *{box-sizing:border-box}
 body{font-family:"Archivo",system-ui,-apple-system,'Segoe UI',Arial,sans-serif;margin:0;background:var(--bg);color:var(--ink);-webkit-font-smoothing:antialiased}
 a{color:var(--accent)}
 header{position:sticky;top:0;z-index:40;background:var(--navy);color:#fff;box-shadow:0 4px 18px rgba(9,20,35,.25)}
 .topbar{max-width:1180px;margin:0 auto;padding:12px 18px;display:flex;align-items:center;gap:16px;flex-wrap:wrap}
 .brand{display:flex;align-items:center;gap:11px;margin-right:auto;text-decoration:none;color:#fff}
 .brand-mark{width:38px;height:38px;border-radius:10px;background:#0878d1;color:#fff;display:grid;place-items:center;font-weight:900;font-size:1.25rem}
 .brand b{font-size:1.12rem;display:block;line-height:1.1}
 .brand span{font-size:.72rem;letter-spacing:.16em;text-transform:uppercase;color:#bcd2e8;font-weight:600}
 nav{display:flex;gap:6px;flex-wrap:wrap;align-items:center}
 nav a{color:#d7e5f2;text-decoration:none;padding:8px 13px;border-radius:999px;font-weight:600;font-size:.9rem}
 nav a:hover{background:rgba(255,255,255,.10);color:#fff}
 nav a.cta{background:#0878d1;color:#fff}
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
 .stat{font-size:1.65rem;font-weight:900;line-height:1;color:var(--accent)}
 /* Premium polish: soft entrance, hover lift, smooth buttons, focus rings */
 @keyframes rise{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
 main>*{animation:rise .45s ease both}
 main>*:nth-child(2){animation-delay:.05s} main>*:nth-child(3){animation-delay:.1s}
 .card{transition:transform .2s ease,box-shadow .2s ease}
 .card:hover{transform:translateY(-2px);box-shadow:0 16px 34px rgba(15,46,77,.16)}
 .hero{background:linear-gradient(135deg,var(--navy) 0%,var(--navy2) 55%,#0878d1 100%);position:relative;overflow:hidden}
 .hero:after{content:"✈";position:absolute;right:18px;bottom:-24px;font-size:9rem;opacity:.08;transform:rotate(-18deg)}
 button,.btn{transition:transform .15s ease,filter .15s ease,box-shadow .15s ease;box-shadow:0 4px 14px rgba(8,120,209,.25)}
 button:hover,.btn:hover{transform:translateY(-1px)}
 button:active,.btn:active{transform:translateY(0) scale(.98)}
 input:focus,select:focus{outline:2px solid var(--accent);outline-offset:1px;border-color:var(--accent)}
 tr{transition:background .15s ease} tbody tr:hover{background:var(--surface2)}
 nav a{transition:background .15s ease,color .15s ease}
 .brand-mark{transition:transform .25s ease} .brand:hover .brand-mark{transform:rotate(-12deg) scale(1.06)}
 /* Premium auth card */
 .auth-card{max-width:460px;margin:28px auto;padding:30px 30px 24px;border-radius:22px;box-shadow:0 20px 44px rgba(15,46,77,.16);position:relative;overflow:hidden}
 .auth-card:before{content:"";position:absolute;top:0;left:0;right:0;height:5px;background:linear-gradient(90deg,var(--navy),var(--accent),#4aa8ec)}
 .auth-badge{display:inline-flex;align-items:center;gap:8px;background:var(--accent-soft);color:var(--navy);font-weight:800;font-size:.82rem;padding:7px 12px;border-radius:999px;margin-bottom:14px}
 .auth-card h2{font-size:1.7rem;margin:0 0 6px}
 .auth-sub{color:var(--muted);margin:0 0 20px;line-height:1.55}
 .auth-card label{font-weight:600}
 .auth-card input{padding:13px 14px;border-radius:12px;font-size:1rem}
 .auth-card button{width:100%;padding:13px;margin-top:6px;font-size:1.02rem;border-radius:12px}
 .auth-switch{text-align:center;color:var(--muted);margin-top:16px}
 .auth-note{text-align:center;color:var(--faint);font-size:.85rem;margin-top:8px;line-height:1.5}
 @media (prefers-reduced-motion:reduce){*{animation:none!important;transition:none!important}}
 .search-card{padding:24px}
 .search-form{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;align-items:end}
 .search-form input,.search-form select{width:100%;margin:4px 0 0}
 .route-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:12px}
 .route-card h3{margin:0 0 4px} .route-card .from-fare{font-size:1.4rem;font-weight:900;color:var(--accent)}
 .trust-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px}
 .faq details{border-bottom:1px solid var(--line2);padding:10px 0} .faq summary{font-weight:700;cursor:pointer}
 @media print{header,.noprint{display:none}}

 /* Edge system: one consistent corner language everywhere */
 :root{--radius:20px}
 body{overflow-x:hidden}
 .card{border-radius:20px}
 .hero{border-radius:24px}
 button,.btn{border-radius:12px}
 nav a{border-radius:999px}
 input,select,textarea{border-radius:12px}
 .brand-mark{border-radius:12px}
 .flash{border-radius:12px}
 .auth-card{border-radius:20px}
 .search-card{border-radius:20px}
 @media (max-width:600px){ main{padding-left:14px;padding-right:14px} }
</style></head><body>
<header><div class="topbar">
 <a class="brand" href="{{ url_for('home') }}"><span class="brand-mark">✈</span><span><b>Horizon Airways</b><span>Airline Reservation System</span></span></a>
 <nav>
 <a href="{{ url_for('home') }}">Main Menu</a>
 <a href="{{ url_for('reservations') }}">Reservations</a>
 <a class="cta" href="{{ url_for('new_reservation') }}">New Reservation</a>
 <a href="{{ url_for('cancel') }}">Cancel Reservation</a>
 <a href="{{ url_for('flights') }}">Flights</a>
 <a href="{{ url_for('flight_status') }}">Flight Status</a>
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
    <div class="card auth-card"><div class="auth-badge">✈ Horizon Airways</div><h2>Create your account</h2>
    <p class="auth-sub">Book flights, view your tickets and manage your reservations.</p><form method="post">
    <label>Full Name<input name="name" required placeholder="e.g. Aarav Sharma"></label>
    <label>Email<input name="email" type="email" required placeholder="you@example.com"></label>
    <label>Password<input name="password" type="password" minlength="6" required placeholder="At least 6 characters"></label>
    <button>Create Account</button></form>
    <p class="auth-switch">Already have an account? <a href="{{ url_for('login') }}">Login</a></p></div>
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
    <div class="card auth-card"><div class="auth-badge">✈ Horizon Airways</div><h2>Welcome back</h2>
    <p class="auth-sub">Login to book flights and view your reservations.</p><form method="post">
    <label>Email<input name="email" type="email" required placeholder="you@example.com"></label>
    <label>Password<input name="password" type="password" required placeholder="Your password"></label>
    <button>Login</button></form>
    <p class="auth-switch">Don't have an account? <a href="{{ url_for('register') }}">Sign Up</a></p>
    <p class="auth-note">Logging in with the owner's email signs you in as admin automatically.</p></div>
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
    cities = ["Delhi", "Mumbai", "Lucknow", "Bengaluru", "Chennai", "Hyderabad", "Kolkata", "Ahmedabad", "Jaipur", "Patna", "Varanasi", "Pune", "Goa", "Kochi", "Guwahati", "Srinagar", "Dehradun", "Chandigarh", "Bhubaneswar", "Indore", "Nagpur", "Ranchi", "Raipur", "Amritsar", "Coimbatore", "Visakhapatnam", "Tirupati", "Agartala", "Port Blair"]
    city_opts = "".join(f'<option>{c}</option>' for c in cities)
    routes = [
        ("Delhi", "Mumbai", "HZ101", "06:00 - 08:10", 4500),
        ("Delhi", "Bengaluru", "HZ202", "09:00 - 11:45", 5200),
        ("Mumbai", "Goa", "HZ303", "13:00 - 14:15", 3200),
        ("Lucknow", "Delhi", "HZ404", "07:30 - 08:40", 2800),
        ("Lucknow", "Bengaluru", "HZ518", "16:55 - 20:30", 7800),
        ("Delhi", "Hyderabad", "HZ504", "09:45 - 12:35", 5200),
    ]
    route_cards = "".join(
        f'<div class="card route-card"><h3>{a} to {b}</h3><p style="color:var(--muted)">{fno} - {tm} - Non-stop</p>'
        f'<p><span style="color:var(--muted);font-size:.85rem">Starting from</span><br>'
        f'<span class="from-fare">Rs.{fare:,}</span></p>'
        f'<p><a class="btn" href="{url_for("flights")}?from_city={a}&to_city={b}">View Flights</a> '
        f'<a class="btn ghost" href="{url_for("new_reservation")}?flight={fno}">Book Now</a></p></div>'
        for a, b, fno, tm, fare in routes
    )
    body = f"""
    <div class="hero"><span class="hero-tag">Horizon Airways</span>
    <h1>Fly smart. Book in under a minute.</h1>
    <p>Search a route, compare fares across classes, and get a confirmed PNR with an e-ticket instantly. When a class fills up, your booking joins the waiting list instead.</p>
    <div class="card search-card" style="margin:20px 0 0;color:var(--ink)">
    <form class="search-form" method="get" action="{url_for('flights')}">
    <label>From<select name="from_city"><option value="">Anywhere</option>{city_opts}</select></label>
    <label>To<select name="to_city"><option value="">Anywhere</option>{city_opts}</select></label>
    <label>Travel Date<input type="date" name="travel_date"></label>
    <label>Class<select name="travel_class"><option>Economy</option><option>Business</option><option>First</option></select></label>
    <label>Sort<select name="sort"><option value="cheapest">Cheapest first</option><option value="earliest">Earliest departure</option><option value="fastest">Fastest</option></select></label>
    <button>Search Flights</button></form></div></div>
    <h2 style="margin:6px 0 12px">Popular routes</h2>
    <div class="route-grid">{route_cards}</div>
    <p style="color:var(--muted);font-size:.9rem">Note: routes added from the All-India airport list use demo fares and demo timings, not live airline fares.</p>
    <div class="grid">{cards}</div>
    <div class="card"><h3>Why book with Horizon?</h3><div class="trust-grid">
    <div><b>Instant PNR</b><p style="color:var(--muted)">Every booking gets a 5-character PNR and a printable e-ticket right away.</p></div>
    <div><b>Transparent fares</b><p style="color:var(--muted)">The fare on your ticket is exactly the flight's class fare - no extra charges are added in this app.</p></div>
    <div><b>Easy cancellation</b><p style="color:var(--muted)">Wait-listed bookings get a full refund. Confirmed bookings within 1 day of travel get 90%, otherwise a full refund.</p></div>
    <div><b>Seat &amp; meal choice</b><p style="color:var(--muted)">Pick Window/Aisle and Veg/Non-veg while booking, just like a full booking flow.</p></div>
    </div></div>
    <div class="card faq"><h3>Frequently asked questions</h3>
    <details><summary>What happens if my class is full?</summary><p style="color:var(--muted)">Your booking becomes Wait-listed. If a Confirmed Economy passenger cancels, the earliest Wait-listed passenger is promoted to Confirmed automatically.</p></details>
    <details><summary>How do refunds work?</summary><p style="color:var(--muted)">Wait-listed: full refund. Confirmed and travelling within 1 day: 90% refund. Confirmed otherwise: full refund. A refund receipt is generated on cancellation.</p></details>
    <details><summary>How do I become an admin?</summary><p style="color:var(--muted)">There is no separate admin signup. Logging in with the owner's email automatically signs you in as an admin, with access to Reports and flight management.</p></details>
    <details><summary>Is this a real airline booking?</summary><p style="color:var(--muted)">No - this is a student project demo. Bookings are stored in this app's database and are not real airline tickets. Real flight inventory and payment would need a supported airline provider integration.</p></details>
    </div>
    <div class="card"><h3>Quick links</h3>
    <p><a class="btn" href="{url_for('new_reservation')}">Make a Reservation</a>
    <a class="btn ghost" href="{url_for('cancel')}">Cancel a Reservation</a>
    <a class="btn ghost" href="{url_for('flight_status')}">Check Flight Status</a>
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
    _from = request.args.get("from_city", "")
    _to = request.args.get("to_city", "")
    _sort = request.args.get("sort", "cheapest")
    _travel_date = request.args.get("travel_date", "")
    _tclass = request.args.get("travel_class", "")
    raw = db.execute(
        """SELECT f.*, s.origin, s.destination FROM flights f
           JOIN sectors s ON s.sector_id=f.sector_id ORDER BY f.flight_no"""
    ).fetchall()
    def _mins(t):
        try:
            h, m = t.split(":"); return int(h) * 60 + int(m)
        except Exception:
            return 0
    rows = []
    for r in raw:
        if _from and r["origin"] != _from:
            continue
        if _to and r["destination"] != _to:
            continue
        d = dict(r)
        fares = [f for f in (r["first_fare"], r["business_fare"], r["economy_fare"]) if f]
        d["from_fare"] = min(fares) if fares else 0
        dur = (_mins(r["arr_time"]) - _mins(r["dep_time"])) % (24 * 60)
        d["duration"] = f"{dur // 60}h {dur % 60:02d}m"
        d["_dur"] = dur
        rows.append(d)
    if _sort == "earliest":
        rows.sort(key=lambda r: r["dep_time"])
    elif _sort == "fastest":
        rows.sort(key=lambda r: r["_dur"])
    else:
        rows.sort(key=lambda r: r["from_fare"])
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
    # 1-stop connecting options (only when no direct flight matches the search)
    _conn_html = ""
    if _from and _to and not rows:
        _class_key = {"Economy": "economy_fare", "Business": "business_fare", "First": "first_fare"}.get(_tclass, "economy_fare")
        _conn = []
        for _leg1 in raw:
            if _leg1["origin"] != _from:
                continue
            for _leg2 in raw:
                if _leg2["origin"] != _leg1["destination"] or _leg2["destination"] != _to:
                    continue
                if not _leg1[_class_key] or not _leg2[_class_key]:
                    continue
                _lay = (_mins(_leg2["dep_time"]) - _mins(_leg1["arr_time"])) % (24 * 60)
                _overnight = _mins(_leg2["dep_time"]) < _mins(_leg1["arr_time"])
                _journey = (_mins(_leg2["arr_time"]) - _mins(_leg1["dep_time"])) % (24 * 60)
                if _journey == 0:
                    _journey = 24 * 60
                _conn.append((_leg1, _leg2, _leg1[_class_key] + _leg2[_class_key], _lay, _overnight, _journey))
        _conn.sort(key=lambda c: c[2])
        if _conn:
            _cards = ""
            for _l1, _l2, _total, _lay, _over, _jour in _conn:
                _cards += (
                    f'<div class="card"><h3>1-stop via {_l1["destination"]} &mdash; {_from} to {_to}</h3>'
                    f'<p><b>Leg 1:</b> {_l1["flight_no"]} &mdash; {_l1["origin"]} to {_l1["destination"]} ({_l1["dep_time"]} &ndash; {_l1["arr_time"]})<br>'
                    f'<b>Leg 2:</b> {_l2["flight_no"]} &mdash; {_l2["origin"]} to {_l2["destination"]} ({_l2["dep_time"]} &ndash; {_l2["arr_time"]})</p>'
                    f'<p><b>Total fare ({_tclass or "Economy"}):</b> Rs.{_total:,} &nbsp; <b>Journey:</b> {_jour // 60}h {_jour % 60:02d}m &nbsp; <b>Layover:</b> {_lay // 60}h {_lay % 60:02d}m'
                    + (" (overnight)" if _over else "") + '</p>'
                    f'<p><a class="btn" href="{url_for("new_reservation")}?flight={_l1["flight_no"]}&travel_class={_tclass or "Economy"}">Book Leg 1 ({_l1["flight_no"]})</a> '
                    f'<a class="btn ghost" href="{url_for("new_reservation")}?flight={_l2["flight_no"]}&travel_class={_tclass or "Economy"}">Book Leg 2 ({_l2["flight_no"]})</a></p>'
                    f'<p style="color:var(--muted);font-size:.9rem">Note: each reservation in this app is for one flight, so a connecting journey needs 2 separate bookings (2 PNRs) &mdash; one per leg.</p></div>'
                )
            _conn_html = f'<h2 style="margin:6px 0 12px">No direct flight &mdash; 1-stop options</h2>{_cards}'
    _cities = ["Delhi", "Mumbai", "Lucknow", "Bengaluru", "Chennai", "Hyderabad", "Kolkata", "Ahmedabad", "Jaipur", "Patna", "Varanasi", "Pune", "Goa", "Kochi", "Guwahati", "Srinagar", "Dehradun", "Chandigarh", "Bhubaneswar", "Indore", "Nagpur", "Ranchi", "Raipur", "Amritsar", "Coimbatore", "Visakhapatnam", "Tirupati", "Agartala", "Port Blair"]
    def _copts(sel):
        return "".join(f'<option{" selected" if c == sel else ""}>{c}</option>' for c in _cities)
    _note = ""
    if _from or _to:
        _note = f'<p style="color:var(--muted)">Showing flights {_from or "anywhere"} to {_to or "anywhere"}'
        if _travel_date:
            _note += f' on {_travel_date}'
        if _tclass:
            _note += f' - { _tclass} class fares highlighted in the fare columns below'
        _note += '.</p>'
    _empty = "" if rows else ('<p><b>No direct flights on this route yet.</b> Try another route, or <a href="/flights">view all flights</a>.</p>' if not _conn_html else "")
    body = f"""
    <div class="card"><h2>Flights</h2>
    <form class="search-form" method="get" action="{url_for('flights')}">
    <label>From<select name="from_city"><option value="">Anywhere</option>{_copts(_from)}</select></label>
    <label>To<select name="to_city"><option value="">Anywhere</option>{_copts(_to)}</select></label>
    <label>Sort<select name="sort"><option value="cheapest"{" selected" if _sort == "cheapest" else ""}>Cheapest first</option><option value="earliest"{" selected" if _sort == "earliest" else ""}>Earliest departure</option><option value="fastest"{" selected" if _sort == "fastest" else ""}>Fastest</option></select></label>
    <button>Search</button></form>
    {_note}{_empty}
    {_conn_html}
    {table_html(rows, [('flight_no','Flight No.'),('sector_id','Route'),('origin','From'),('destination','To'),('dep_time','Departure'),('arr_time','Arrival'),('duration','Duration'),('from_fare','From Fare (Rs.)'),('first_fare','First Fare (Rs.)'),('business_fare','Business Fare (Rs.)'),('economy_fare','Economy Fare (Rs.)')])}
    <p style="color:var(--muted);font-size:.9rem">Fare comparison: First, Business and Economy fares are shown side by side for every flight. A fare of 0 means that class is not offered on that flight. All flights shown are non-stop.</p>
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


@app.route("/flight-status")
def flight_status():
    db = get_db()
    fno = request.args.get("flight_no", "").strip().upper()
    result = ""
    if fno:
        r = db.execute(
            """SELECT f.*, s.origin, s.destination FROM flights f
               JOIN sectors s ON s.sector_id=f.sector_id WHERE f.flight_no=?""", (fno,)
        ).fetchone()
        if r:
            fares = [f for f in (r["first_fare"], r["business_fare"], r["economy_fare"]) if f]
            result = f"""<div class="card"><h3>{r["flight_no"]} - {r["origin"]} to {r["destination"]}</h3>
            <p><b>Departure:</b> {r["dep_time"]} &nbsp; <b>Arrival:</b> {r["arr_time"]} &nbsp; <b>Aircraft:</b> {r["aircraft_id"]}</p>
            <p><b>Operates:</b> {r["weekday1"]} and {r["weekday2"]} &nbsp; <b>Starting from:</b> Rs.{min(fares) if fares else 0:,}</p>
            <p style="color:var(--muted)">Status: Scheduled (as per this app's timetable). This is a demo project, not a live airline feed, so real-time delays are not shown.</p>
            <p><a class="btn" href="{url_for('new_reservation')}?flight={r["flight_no"]}">Book This Flight</a></p></div>"""
        else:
            result = f'<div class="card"><p><b>No flight found for {fno}.</b> Try HZ101, HZ202, HZ303 or HZ404.</p></div>'
    body = f"""<div class="card"><h2>Flight Status</h2>
    <form class="search-form" method="get" action="{url_for('flight_status')}">
    <label>Flight Number<input name="flight_no" value="{fno}" placeholder="e.g. HZ101" required></label>
    <button>Check Status</button></form></div>{result}"""
    return page("Flight Status", body)


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
    _sel_flight = request.args.get("flight", "")
    _sel_class = request.args.get("travel_class", "")
    flight_opts = "".join(
        f'<option value="{r["flight_no"]}"{" selected" if r["flight_no"] == _sel_flight else ""}>{r["flight_no"]} — {r["origin"]} → {r["destination"]} ({r["dep_time"]})</option>'
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
    <label>Class<select name="travel_class"><option{" selected" if _sel_class == "Economy" else ""}>Economy</option><option{" selected" if _sel_class == "Business" else ""}>Business</option><option{" selected" if _sel_class == "First" else ""}>First</option></select></label>
    <label>Seat Preference<select name="seat_pref"><option>Window</option><option>Aisle</option></select></label>
    <label>Meal Preference<select name="meal_pref"><option>Veg</option><option>Non-veg</option></select></label>
    <label>Special Service Request<input name="ssr" placeholder="e.g. Wheelchair"></label>
    <p><button>Save Reservation &amp; Print Ticket</button></p>
    <p style="font-size:13px;color:#555">Transparent pricing: the fare on your ticket is exactly the flight's class fare - no extra charges are added in this app. Fare is taken automatically from the flight's class fare.
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
