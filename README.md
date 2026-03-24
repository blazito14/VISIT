# Student Visit Scheduler

This folder contains everything needed to run the scheduler with as little setup as possible on Windows.

## Files included

- `visit_scheduler.py` — the main scheduler script
- `scheduler_gui.py` — a simple windowed launcher
- `setup.bat` — creates `.venv` and installs dependencies
- `run_scheduler.bat` — opens the GUI
- `requirements.txt` — Python packages needed
- `.gitignore` — ignores the virtual environment and outputs

## Easiest setup for users

1. Install Python 3
2. Download this project from GitLab
3. Double-click `setup.bat`
4. Double-click `run_scheduler.bat`
5. Choose the two CSV files and click **Run Scheduler**

## Recommended GitLab structure

Put these files in the root of your repo:

- `visit_scheduler.py`
- `scheduler_gui.py`
- `setup.bat`
- `run_scheduler.bat`
- `requirements.txt`
- `README.md`

## Notes

- This setup is aimed at Windows users.
- The GUI uses `tkinter`, which is normally included with standard Python installs.
- Output files will be written to the output folder chosen in the GUI.
