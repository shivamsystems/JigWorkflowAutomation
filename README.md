# 🛠️ Jig Workflow Automation

![Python](https://img.shields.io/badge/Python-3.9+-blue.svg)
![Windows](https://img.shields.io/badge/Platform-Windows-0078D6.svg)
![Status](https://img.shields.io/badge/Status-Active-success.svg)

A background automation tool for Mechanical/Manufacturing engineers that tracks SolidWorks design sessions, 3D printing material usage, and daily part handovers. No more forgetting to log your hours or forgetting what parts you handed off to production at 5:15 PM.

## 🎯 The Problem
As a jig & fixture designer, my daily workflow involves jumping between SolidWorks, Bambu Studio, and physical handovers to the production floor. Tracking billable hours, filament consumption, and part handovers was a manual, error-prone process that often got skipped.

## 💡 The Solution
A Python-based background daemon that uses process monitoring and event-driven architecture to detect when I open/close software, automatically logging time and prompting for data entry precisely when needed.

## ✨ Features
* **SolidWorks Tracking**: Detects when SolidWorks opens, prompts for project details, and calculates exact session duration on close.
* **3D Printing Logs**: Prompts for filament usage and print results when Bambu Studio/Orcaslicer is closed.
* **5:15 PM Handover System**: Daily reminder with a built-in state machine. If skipped, it reminds you every 15 minutes until 6:00 PM. Missed days are caught on the next boot.
* **Auto-Folder Generation**: Automatically creates standardized project directories (CAD, STL, Prints, etc.).
* **System Tray Integration**: Runs silently in the background. Minimizes to tray on Windows startup.
* **Thread-Safe Excel Logging**: Writes to Excel with retry logic to prevent data corruption if the file is open.

## 🔄 Workflow Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                  JIG WORKFLOW AUTOMATION                    │
├──────────────────────┬──────────────────────────────────────┤
│                      │                                      │
│   SOLIDWORKS FLOW    │       3D PRINTING FLOW               │
│                      │                                      │
│  SW Opens ──────────►│  BambuStudio Closes ──────────────►  │
│       │              │         │                            │
│  ┌────▼─────┐        │   ┌─────▼──────────┐                 │
│  │ Project  │        │   │ Print &        │                 │
│  │ Details  │        │   │ Filament Log   │                 │
│  └────┬─────┘        │   └─────┬──────────┘                 │
│       │              │         │                            │
│  SW Closes ────────► │         │                            │
│       │              │         │                            │
│  ┌────▼─────────┐    │         │                            │
│  │ Session End  │    │         │                            │
│  │ Summary      │    │         │                            │
│  └────┬─────────┘    │         │                            │
├───────┴──────────────┴─────────┴────────────────────────────┤
│                  5:15 PM SCHEDULER                          │
│  • Triggers daily handover log                              │
│  • Smart skip logic (Reminds every 15m, Deadline at 6PM)    │
│  • Next-day startup check for missed logs                   │
├──────────────────────────────────────────────────────────────┤
│                  EXCEL LOGGER (Thread-Safe)                 │
│  • SolidWorks Projects Sheet                                │
│  • Print & Filament Log Sheet                               │
│  • Daily Handovers Sheet                                    │
└──────────────────────────────────────────────────────────────┘
```
## 🚀 Installation & Setup
Prerequisites
Windows 10/11
Python 3.9 or higher (Ensure "Add Python to PATH" is checked during installation)
Step-by-Step
Clone the repository:

```

git clone https://github.com/shivamsystems/JigWorkflowAutomation.git
cd JigWorkflowAutomation
```
Run the installer (sets up virtual environment and dependencies):

```

install.bat
```
(Optional) Auto-Start with Windows:
Right-click add_to_startup.bat and select "Run as Administrator". This creates a Windows Task Scheduler entry that runs the tool silently in your system tray every time you log in.

## 🧠 Technical Deep Dive
This project avoids the standard "top-down" script approach and instead uses Event-Driven Architecture.

* **The Watchman (psutil + root.after)**: Uses process ID (PID) set theory (CurrentList - PreviousList = JustOpened) to detect software state changes without polling APIs.
* **The Scheduler (datetime state machine)**: Instead of blocking time.sleep() loops (which freeze Tkinter UIs), the clock checker runs every 30 seconds, evaluating the time against the 5:15 PM target and 6:00 PM deadline, tracking the application's handover_status (Waiting -> Pending -> Done).
* **The Memory (openpyxl + Threading)**: Implements a threading.Lock with a 3-retry mechanism to prevent PermissionError crashes when attempting to write to the Excel file while the user has it open.
* **The Ghost (pythonw.exe + Task Scheduler)**: Uses a VBS wrapper to launch Python without a console window, keeping the desktop clean while the system tray icon provides UI access.
## 📂 Project Structure
```

JigWorkflowAutomation/
├── jig_workflow.py          # Core application logic, UI, and monitoring
├── requirements.txt         # Python dependencies
├── install.bat              # One-click environment setup
├── run_silent.vbs           # Launches tool without a console window
├── add_to_startup.bat       # Registers tool in Windows Task Scheduler
├── remove_startup.bat       # Removes from Task Scheduler
├── .gitignore               # Ignores logs, venv, and generated data
└── README.md                # You are here
```
Note: settings.json, Jig_Workflow_Log.xlsx, and jig_workflow.log are auto-generated on first run and safely ignored by Git.

