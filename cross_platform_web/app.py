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
 body{font-family:"Archivo",system-ui,-apple-system,'Segoe UI',Arial,sans-serif;margin:0;background:transparent;color:var(--ink);-webkit-font-smoothing:antialiased}html{background:var(--bg)}body{background:transparent}body::before{content:"";position:fixed;inset:0;z-index:-1;pointer-events:none;background-image:url("data:image/gif;base64,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l116+SWYYYo5JpllmnkmmmmquSabbbr5JpxxyjknnXXaeSeeeeq5J599+vknoIEKOiihhRp6KKKJKrooo406+iikkUo6KaWVWnoppplquimnnXr6KaihijoqqaWaeiqqqaq6KqutuvoqrLHKOiuttdp6K6656rorr736+iuwwQo7LLHFGnsssskquyyzzTr7LLTRSjsttdVaey222Wq7LbfdevstuOGKOy655Zp7Lrrpqrsuu+26+y688co7L7312nsvvvnquy+//fr7L8ABCzwwwQUbfDDCCSu8MMMNO/wwxBFLPDHFFVt8Mca5BgQAIfkEAQ4AgAAsAAAPAIACHQGG////1ur61On60+n50uj50ej50ef50Of5z+f5z+f4z+b4zub4zeb4zeX4zOX4y+X4y+T4y+T3yuT3yeT3yeP3yOP3x+P3x+L3xuL3xuL2xeL2xeH2xOH2w+H2w+D2wuD2wuD1wd/1wN/1v971vt71vd30vN30vNz0u9z0utz0utv0udv0uNr0uNrzt9rzt9nzttnztdnztdjztNjzs9jzs9fzstfyv8+/sdbysNbysNXyr9XyrtXyrtTyrdTyrdTxrNPxq9PxqtLxqdLxqdHxqNHxp9Hwp9DwptDwpc/wpM/w/808vM2/tMm9pM7wrsa8q8S7o87wos7wos3voc3voM3voMzvn8zvnsvvncvvncrvnMrum8rum8numsnumcjumMjul8ful8ftlsftlsbtlcbtlMbtlMXtk8XtksTtkcTskcPskMPsj8Psj8LsjsLsjcLsjcHsjMHsjMHrjMDri8DrisDrir/rib/riL/riL7rh77reKrNLXilFFqMCi1QCP8AAQkcSLCgwYMIEypcyLChw4cQI0pcCKCixYoG62jcqHGix48gQ4ocSbKkyZMoU6pcybKly5cwYxa8SJOmQI44ccrcybOnz59AgwodSrSoUYU1kwLImLPp0adQo0qdSrWq1asoldZESKer169fsYodS7as2bNo0Wq1uXCO27dw4aadS7eu3bt48w5ce9GhnL+AAwfWS7iw4cOIE4fkaxGi4MeAFUueTLmyZbGMMUaMw7mzZ8+XQ4seTbo0ycxLJ35e3dm069ewY19G7ZE1a9m4c+veTZb2RzjAgwsXzru48ePIY/oG+aa58+fPk0u3+6S69evYnwTNzt369NmZRUL/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u9z0utz0utv0udv0uNr0uNrzt9rzt9nzttnztdnztdjztNjzs9jzs9fzstfyv8+/sdbysNbysNXyr9XyrtXyrtTyrdTyrdTxrNPxq9PxqtLxqdLxqdHxqNHxp9Hwp9DwptDwpc/wpM/w/808usy+rsW8pM7wo87wos7woc3voM3voMzvn8zvnsvvncvvncrvnMrum8rum8numsnumcjumMjul8ful8ftlsftlsbtlcbtlMbtlMXtk8XtksTtkcTskcPskMPsj8Psj8LsjsLsjcLsjcHsjMHsjMHrjMDri8DrisDrir/rib/riL/riL7rh77reKrNLXilFFqMCi1QAAAACP8A/wgcSLCgwYMIEypcyLChw4cQI0qcSLGixYsYM2rcyLGjx48gMz4ZSbKkyZAoU6pcybKly5cwY8qcSbMmRZM4c+a0ybOnz59AgwodSrQoR51Ik+I0yrSp06dQo0qdulKp1askqWrdyrWr169gN2IdizWs2bNo06pdC5Os27Js48qdS7eu2bd44drdy7ev378q8wrWy3POHMCIEyteLHOwY8I05UhmTLmy5csRH2u2yjOO5ziYQ4serXhzQ9M0P3smzbq167SDL8aeCae27de4c+t2mrdjb5m2g+8eTrx4TLwpkcMMLty48+fQMb5tOf3lm+vYr0ffzr07QrfHyb7/dEO+fHnv6NM7F0+TPUvz8NXLn+96rE/77+Gbpy8agP///hmkH3n8kYbffZCt1MaCDP7B4IMF/gXghBM+aKGFEVp2FYUTDtTEhyCG2ARIB7504YltZEgXhxwShOKFKi6WFIsUeijijR9+dNVMbPTo448/xrgWjR0eBOSRPgr5l05EApgQjjj6xhmPSB6ppFlN/qdQlVZeaVdOWQbIEJQ3HjWlTGukqeaaa3rpVZgAMMTmnGq6OReYYT5EpohmKkUTnXTaqRWcDgE6p6Bs4UQoRHuG2CdSfxraJqJRLVqopHVSepZJlkrUKIhi+TmTGqSWaqqpmjrVqUOntlpqqmBx/5qnRZ/mqNGZMbnqKqxFrfqQrq3y2lVJvk5U64giiSpTGsw266yzwg5VbEPPVttstFoRO2tGx96qLEzWWostUNM6FG6140qlbZYbdestpDOhIe+89NKbrkzlZlTvvvPeS9S6TXbkbrLwxsQvv/62lK9HB++b8E8kLUzRwLIlFW/D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");background-size:cover;background-position:center;background-repeat:no-repeat;opacity:.55}@media (prefers-color-scheme:dark){body::before{opacity:.28}}.card{background:color-mix(in srgb, var(--surface) 94%, transparent)}
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
