# ✈️ Airline Reservation System

A web and desktop application that manages airline reservations — book flights, manage sectors and aircraft, schedule flights, print tickets, cancel reservations with refund calculation, and view daily reports.

The cross-platform web version is built using **Python + Flask + SQLite**, and the original desktop version is built using **Visual Basic .NET WinForms + Microsoft Access**.

---

## 🌐 Live Demo

🔗 **Try the application here:**
https://rishabh4630.github.io/Airline-Reservation-System/

> This is a browser demo of the web version, hosted on GitHub Pages — it works for anyone. Booking in the demo needs a free Login / Sign Up (accounts live only in that browser, and demo bookings reset on reload). For the full version with a saved database, run the web version below.

---

## 📌 Project Overview

This project manages airline reservations for **Horizon Airways** through a simple interface for flight-related operations.

This repository contains **two versions** of the same project:

| Version | Technology | Runs on |
|---|---|---|
| **Cross-Platform Web Version** — [`cross_platform_web/`](cross_platform_web/) | Python + Flask + SQLite | **Windows, macOS and Linux** |
| Original Desktop Version — files in the root of this repo | Visual Basic .NET WinForms + MS Access | Windows only |

> **Using a Mac?** Use the web version in `cross_platform_web/` — the original VB.NET WinForms app cannot run on macOS because WinForms and MS Access are Windows-only technologies.

Users can manage flights and schedules, book a reservation and get a PNR, print a ticket, cancel a reservation and get a refund receipt, and view confirmed, waiting-list and daily-collection reports.

---

## ✨ Features

- ✈️ Flight Management
- 🗓️ Flight Scheduling
- 🧑‍🤝‍🧑 Passenger Reservation Management
- ❌ Cancel Reservations
- 🎫 Ticket Printing
- 🧾 Refund Receipt Printing
- 🗺️ Sector Management
- 🛩️ Aircraft Management
- 📊 Reports:
  - Confirmed Passenger List
  - Overbooking / Waiting List
  - Daily Collections Report
- 💾 Database Integration
- ✅ Automatic fare calculation from the flight's class fare
- 🪑 Seat-availability check — if a class is full, the reservation becomes **Wait-listed**
- 🔁 Cancelling a Confirmed Economy seat promotes the earliest Wait-listed passenger to Confirmed
- 🔐 Login / Sign Up — customers book and manage only their own reservations; logging in with the owner's email automatically signs you in as an admin (no separate admin login page)
- 🔍 Home search, popular routes and flight filters — search by route/date/class, sort by cheapest / earliest / fastest, and compare First / Business / Economy fares side by side
- 🛫 Flight Status — check a flight's timetable, aircraft and starting fare by flight number
- 🔀 1-stop connecting flights — when no direct flight exists (e.g. Varanasi → Bengaluru), the search shows a via-Delhi option with total fare, journey time and layover (booked as 2 separate bookings, one per leg)
- **All-India routes** — 30 airports and 86 demo flights across India (see cross_platform_web README for the full city list); new routes use demo fares/timings, not live airline data.
- **Premium 4K video background** — a cinematic 4K airplane-window video (cloudscape from above, 3840x2160, 60 sec loop, muted) plays as the page wallpaper, with a blue gradient overlay so text stays readable; the vector flight scene (route arc + boarding pass) stays behind it as fallback while the video buffers, and the video is hidden for users who prefer reduced motion
- 🎨 Premium blue-white design with dark-mode support, FAQ and transparent fares (the ticket fare is exactly the flight's class fare)

---

## 📂 Project Structure

```text
Airline-Reservation-System/
│
├── cross_platform_web/
│   ├── app.py
│   ├── README.md
│   ├── requirements.txt
│   ├── run_windows.bat
│   ├── run_mac_linux.sh
│   └── .gitignore
│
├── HorizonApp.sln
├── HorizonApp.vbproj
├── MainMenu.vb
├── Flights.vb
├── ScheduledFlighs.vb
├── Sector.vb
├── Aircraft.vb
├── Reservations.vb
├── CancelReservation.vb
├── PrintTicket.vb
├── PrintRefundReceipt.vb
├── *.resx
├── Database/
│   └── Airline.mdb
└── README.md
```

---

## 🛠️ Technologies Used

- Python
- Flask
- SQLite
- HTML / CSS
- Visual Basic .NET
- .NET Framework
- Visual Studio
- Microsoft Access

---

## 📊 Reservation Details

| Field | Description |
|----------|-------------|
| PNR No. | 5-character booking reference, generated automatically |
| Passenger Name | First name and last name |
| Age | Passenger age |
| Gender | Male / Female / Other |
| Class | First / Business / Economy |
| Flight No. | e.g. HZ101 |
| Travel Date | Date of travel |
| Fare Amount | Taken automatically from the flight's class fare |
| Seat Preference | Window / Aisle |
| Meal Preference | Veg / Non-veg |
| Special Service Request | e.g. Wheelchair |
| Status | Confirmed / Wait-listed / Cancelled |

---

## 🧠 How It Works

### Seat Availability

A reservation is **Confirmed** only if a seat is available in the chosen class on that flight and date. If the class is full, the reservation becomes **Wait-listed**.

### Refund Rule

The refund rule is the same as in the original application:

| Situation | Refund |
|----------|--------|
| Wait-listed / Overbooked | Full refund |
| Confirmed, travel within 1 day | 90% refund |
| Confirmed, otherwise | Full refund |

On cancelling a Confirmed Economy seat, the earliest Wait-listed passenger on that flight and date is promoted to Confirmed.

### Sample Data

Sample data is created automatically on the first run of the web version, so the app works immediately — for example:

| Flight | Route | Departure — Arrival | Economy Fare |
|----------|-------|---------------------|--------------|
| HZ101 | Delhi → Mumbai | 06:00 — 08:10 | $4500 |
| HZ202 | Delhi → Bengaluru | 09:00 — 11:45 | $5200 |
| HZ303 | Mumbai → Goa | 13:00 — 14:15 | $3200 |
| HZ404 | Lucknow → Delhi | 07:30 — 08:40 | $2800 |

---

## ⚙️ Installation

### 1️⃣ Clone the Repository

```bash
git clone https://github.com/rishabh4630/Airline-Reservation-System.git
```

### 2️⃣ Navigate to the Project Folder

```bash
cd Airline-Reservation-System
```

### 3️⃣ Web Version — Navigate to the Web Folder

```bash
cd cross_platform_web
```

### 4️⃣ Create a Virtual Environment

```bash
python3 -m venv venv
```

### 5️⃣ Activate the Virtual Environment

#### Windows

```bash
venv\Scripts\activate
```

#### macOS/Linux

```bash
source venv/bin/activate
```

### 6️⃣ Install Dependencies

```bash
pip install -r requirements.txt
```

### Desktop Version — Windows Only

1. Open `HorizonApp.sln` in Visual Studio on a compatible Windows environment.
2. Restore any required project dependencies.
3. Check the database configuration and connection settings.
4. Build and run the project.

> Note: the original desktop project references `CommonClasses.dll` from outside this repository, so it may not build from a fresh clone without that library. The web version has no such dependency — it is fully standalone.

---

## 📦 Requirements

Create a **requirements.txt** file with the following content:

```text
Flask>=3.0
```

The web version's `requirements.txt` is already included in `cross_platform_web/`.

---

## ▶️ Run the Application

### Web Version

#### Windows

```bash
run_windows.bat
```

Or manually:

```bash
py -m pip install -r requirements.txt
py app.py
```

#### macOS / Linux

```bash
chmod +x run_mac_linux.sh
./run_mac_linux.sh
```

Or manually:

```bash
python3 -m pip install -r requirements.txt
python3 app.py
```

After running the command, open your browser and visit:

```text
http://127.0.0.1:5000
```

> Requirements: Python 3.10 or newer — download from https://www.python.org/downloads/ (on Windows, tick "Add Python to PATH" during install).

---

## 📈 Booking Workflow

1. Choose a flight and travel date
2. Enter passenger details
3. Fare is calculated automatically from the flight's class fare
4. Seat availability is checked — Confirmed, or Wait-listed if the class is full
5. A 5-character PNR is generated
6. Print the e-ticket for the PNR
7. Cancel by PNR when needed — refund is calculated and a refund receipt is printed
8. View Confirmed, Waiting List and Daily Collections reports

---

## 📷 Application Preview

Add screenshots here after running the application.

```text
assets/
├── dashboard.png
├── reservation.png
├── ticket.png
```

Example:

```markdown
![Dashboard](assets/dashboard.png)

![Reservation](assets/reservation.png)

![Ticket](assets/ticket.png)
```

---

## 🚀 Future Improvements

- Add online payment integration
- Send e-tickets by email
- Add seat-map selection
- Deploy using Docker
- Add real flight-data and live status integration (currently the timetable is demo data)

---

## 🤝 Contributing

Contributions are welcome!

1. Fork this repository

2. Create a new branch

```bash
git checkout -b feature-name
```

3. Commit your changes

```bash
git commit -m "Added new feature"
```

4. Push to GitHub

```bash
git push origin feature-name
```

5. Create a Pull Request

---

## 👨‍💻 Author

**Rishabh Singh Rajput**

B.Tech Computer Science & Engineering Student

Interested in:

- Web Development
- Problem-Solving
- Python Development
- Software Engineering

### GitHub

https://github.com/rishabh4630

---

## ⭐ Show Your Support

If you found this project useful, please consider giving it a ⭐ on GitHub.

It motivates me to build more projects.

---

## 📄 License

This project is intended for educational and learning purposes.

Feel free to use, modify, and share it for learning purposes.
