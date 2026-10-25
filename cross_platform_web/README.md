# Airline Reservation System — Cross-Platform Web Version

This folder is the **Mac + Windows (and Linux)** version of the Airline
Reservation System.

The original project (in the root of this repository) is built with
**Visual Basic .NET WinForms + Microsoft Access (`.mdb`)**. WinForms and
Access only run on Windows, so that version can never run natively on a Mac.

This version keeps the **same features and business rules**, but is rebuilt
with technology that runs everywhere:

| Original (Windows only) | This version (all platforms) |
|---|---|
| Visual Basic .NET WinForms | Python + Flask (web pages in the browser) |
| Microsoft Access `Airline.mdb` / SQL Server | SQLite `airline.db` (single file, no server) |
| Visual Studio on Windows | Any editor, on Windows / macOS / Linux |
| `.exe` output | Open `http://127.0.0.1:5000` in a browser |

## Design

The web interface uses the same premium blue-white look as the GitHub
Pages demo — navy (#063B78) and bright blue (#0878D1) on a light-blue
background, rounded cards, and dark-mode support from your system
setting. The home page has a flight search widget, popular routes,
a Why Horizon section and FAQ.

## Features (same as the original, plus login)

- **Sign Up / Login** — `/register` and `/login`: a customer can book,
  view and cancel only their own reservations (Reservations shows their
  own bookings), and logout. There is no separate admin signup/login
  page: logging in with the owner's email (`OWNER_EMAIL` in `app.py`)
  automatically signs you in as an admin.
- **Profile & Settings** — `/profile` (login required): premium profile
  card with booking stats and recent bookings, edit name, change
  password, theme (System/Light/Dark), default departure city for the
  home search, plus clear-my-bookings and delete-account in the danger
  zone.
- **Admin (owner only)** — an admin can view all reservations, open
  Reports, and add/edit flights, sectors, aircraft and schedules.
- Passwords are stored only as hashes, never in plain text. Booking,
  cancelling and tickets require login; browsing the master tables does
  not, but editing them is admin-only.
- **Main Menu / Dashboard** — flight search (route, date, class, sort),
  popular routes with starting fares, Why Horizon, FAQ, and counts for
  flights, sectors, aircraft, reservations
- **1-stop connecting flights** — when no direct flight exists (e.g. Varanasi → Bengaluru), Flights shows a via-Delhi option with total fare, journey time and layover; each leg is booked separately (2 PNRs)
- **All-India routes** — 30 airports and 86 demo flights covering Delhi, Mumbai, Lucknow, Bengaluru, Chennai, Hyderabad, Kolkata, Ahmedabad, Jaipur, Patna, Varanasi, Pune, Goa, Kochi, Guwahati, Srinagar, Dehradun, Chandigarh, Bhubaneswar, Indore, Nagpur, Ranchi, Raipur, Amritsar, Coimbatore, Visakhapatnam, Tirupati, Agartala and Port Blair; new routes use demo fares/timings (not live airline data), and older databases gain the new routes automatically on the next run
- **Premium 4K video background** — a cinematic 4K airplane-window video (cloudscape from above, 3840x2160, 60 sec loop, muted) plays as the page wallpaper, with a blue gradient overlay so text stays readable; the vector flight scene (route arc + boarding pass) stays behind it as fallback while the video buffers, and the video is hidden for users who prefer reduced motion
- **Flight Status** — `/flight-status`: check a flight's timetable,
  aircraft and starting fare by flight number (schedule-based, not a
  live airline feed)
- **Sectors** — add / edit sectors (Sector ID, origin, destination)
- **Aircraft** — add / edit aircraft types and class-wise seat counts
- **Flights** — flight number, sector, aircraft, departure/arrival time,
  weekdays, duration, starting fare, and First / Business / Economy
  fares side by side; filter by route and sort by cheapest / earliest /
  fastest
- **Schedule Flights** — schedule a flight on a date
- **Reservations** — passenger name, age, gender, class, seat preference
  (Window/Aisle), meal preference (Veg/Non-veg), special service request;
  fare is taken automatically from the flight, and seat availability is
  checked — if the class is full the status becomes **Wait-listed**
- **Print Ticket** — printable Horizon Airways ticket for a PNR
- **Cancel Reservation** — get details by PNR and cancel, with the original
  refund rule:
  - Wait-listed / Overbooked → **full refund**
  - Confirmed and travel is within 1 day → **90% refund**
  - Confirmed otherwise → **full refund**
  - Cancelling a Confirmed Economy seat promotes the earliest
    Wait-listed passenger to Confirmed
- **Print Refund Receipt** — printable refund receipt
- **Reports** — Confirmed Passenger List, Overbooking/Waiting List,
  Daily Collections Report

Sample data (sectors, aircraft and flights like HZ101 Delhi → Mumbai)
is created automatically on the first run, so the app works immediately.

## How to Run

### Windows

```bat
run_windows.bat
```

Or manually:

```bat
py -m pip install -r requirements.txt
py app.py
```

### macOS

```bash
chmod +x run_mac_linux.sh
./run_mac_linux.sh
```

Or manually:

```bash
python3 -m pip install -r requirements.txt
python3 app.py
```

### Linux

```bash
chmod +x run_mac_linux.sh
./run_mac_linux.sh
```

Then open **http://127.0.0.1:5000** in your browser (Chrome, Edge, Safari,
Firefox — any browser, on any OS).

> Requirements: Python 3.10 or newer — download from
> https://www.python.org/downloads/ (on Windows, tick
> "Add Python to PATH" during install).

## Files

- `app.py` — the whole application (routes, SQLite database, business rules)
- `requirements.txt` — the only dependency: Flask
- `airline.db` — SQLite database, created automatically on first run
  (delete it to reset the app to the sample data)
- `run_windows.bat` / `run_mac_linux.sh` — one-click launchers

## Author

Rishabh Singh Rajput — cross-platform port of the original VB.NET project,
for educational and learning purposes.
