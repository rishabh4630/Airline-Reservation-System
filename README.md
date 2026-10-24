# Airline Reservation System

## 🔗 Live Demo

**Try it online — no installation needed:**

[Live Demo — Airline Reservation System](https://rishabh4630.github.io/Airline-Reservation-System/)

> This is a browser demo of the web version, hosted on GitHub Pages — it works for anyone, no account or signup needed. You can book flights, print tickets, cancel reservations and view reports right in the browser. Demo data is not saved — for the full version with a saved database, run the web version below.
>
> Note: if the link shows a 404 the first time, enable Pages once in the repo: Settings → Pages → Build and deployment → Source: Deploy from a branch → Branch: gh-pages → Save.

## Project Overview
The Airline Reservation System is a software application designed to manage airline reservations and provide a simple interface for handling flight-related operations.

This repository contains **two versions** of the same project:

| Version | Technology | Runs on |
|---|---|---|
| **Cross-Platform Web Version** — [`cross_platform_web/`](cross_platform_web/) | Python + Flask + SQLite | **Windows, macOS and Linux** |
| Original Desktop Version — files in the root of this repo | Visual Basic .NET WinForms + MS Access | Windows only |

> **Using a Mac?** Use the web version in `cross_platform_web/` — the original
> VB.NET WinForms app cannot run on macOS because WinForms and MS Access are
> Windows-only technologies.

## Features
- Flight Management
- Flight Scheduling
- Passenger Reservation Management
- Cancel Reservations (with refund calculation)
- Ticket Printing
- Refund Receipt Printing
- Sector Management
- Aircraft Management
- Reports: Confirmed Passengers, Waiting List, Daily Collections
- Database Integration

## How to Run — Cross-Platform Web Version (Windows + Mac + Linux)

See [`cross_platform_web/README.md`](cross_platform_web/README.md) for full details.

**Windows:**
```bat
cd cross_platform_web
run_windows.bat
```

**macOS / Linux:**
```bash
cd cross_platform_web
chmod +x run_mac_linux.sh
./run_mac_linux.sh
```

Then open http://127.0.0.1:5000 in your browser.

## How to Run — Original Desktop Version (Windows only)

1. Download or clone this repository.
2. Open `HorizonApp.sln` in Visual Studio on a compatible Windows environment.
3. Restore any required project dependencies.
4. Check the database configuration and connection settings.
5. Build and run the project.

## Technologies Used
- **Web version:** Python, Flask, SQLite
- **Desktop version:** Visual Basic .NET, .NET Framework, Visual Studio, MS Access (see the `Database` folder)

## Project Structure
- `cross_platform_web/` – Cross-platform Python web version (recommended for Mac users)
- `HorizonApp.sln` – Visual Studio solution file (desktop version)
- `HorizonApp.vbproj` – Project configuration (desktop version)
- `.vb` files – Application source code (desktop version)
- `.resx` files – Form resources (desktop version)
- `Database/` – Database-related files (desktop version)

## Installation
Clone the repository using:

```bash
git clone https://github.com/rishabh4630/Airline-Reservation-System.git
```

## Author
Rishabh Singh Rajput

## License
This project is intended for educational and learning purposes.
