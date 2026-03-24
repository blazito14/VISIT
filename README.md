# Student Visit Scheduler

This folder contains everything needed to run the scheduler with as little setup as possible on Windows.

## Files included

- `visit_scheduler.py` — the main scheduler script
- `scheduler_gui.py` — a simple windowed launcher
- `setup.bat` — creates `.venv` and installs dependencies
- `run_scheduler.bat` — opens the GUI
- `requirements.txt` — Python packages needed
- `.gitignore` — ignores the virtual environment and outputs
- `Tutorial.pptx` — Step by Step instructions with pictures
- `README.md` — This file :) 

## Easiest setup for users

1. Copy the Student Response Form located [here](https://docs.google.com/forms/d/12F-sr8z_ELR9Bnltr4g9MHVci3ZakqCK2QatgF3elPU/copy)
2. Copy the Teacher Response form located [here](https://docs.google.com/forms/d/1g8YOdl-VipPWdG_wlhJtn_TcVO_2YgEighP2WUMV2bw/copy)
3. Publish the forms and share the respective links with students and teachers
4. Link each form response to Google Sheets
5. Download each sheet as a .csv
6. Install Python [here](https://www.python.org/downloads/)
7. Download this project from GitLab as a .zip
8. Unzip the folder
9. Double-click `setup.bat`
10. Double-click `run_scheduler.bat`
11. Choose the two CSV files and click **Run Scheduler**
12. Follow the prompts to select each csv and a desired output folder
13. Use the outputs to automatically and manually schedule groups

## Notes

- This setup is aimed at Windows users.
- The GUI uses `tkinter`, which is normally included with standard Python installs.
- Output files will be written to the output folder chosen in the GUI.
