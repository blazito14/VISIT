import os
import subprocess
import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "visit_scheduler.py"

def browse_file(entry_var, title):
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
    )
    if path:
        entry_var.set(path)

def browse_folder(entry_var, title):
    path = filedialog.askdirectory(title=title)
    if path:
        entry_var.set(path)

def run_scheduler():
    student_csv = student_var.get().strip()
    teacher_csv = teacher_var.get().strip()
    out_dir = out_var.get().strip()
    group_size = group_size_var.get().strip()
    visits_per_group = visits_var.get().strip()

    if not student_csv or not os.path.isfile(student_csv):
        messagebox.showerror("Missing file", "Please select a valid student CSV file.")
        return

    if not teacher_csv or not os.path.isfile(teacher_csv):
        messagebox.showerror("Missing file", "Please select a valid teacher CSV file.")
        return

    if not out_dir:
        messagebox.showerror("Missing folder", "Please choose an output folder.")
        return

    try:
        group_size_int = int(group_size)
        visits_int = int(visits_per_group)
        if group_size_int < 1 or visits_int < 1:
            raise ValueError
    except ValueError:
        messagebox.showerror("Invalid settings", "Group size and visits per group must be positive whole numbers.")
        return

    os.makedirs(out_dir, exist_ok=True)

    cmd = [
        sys.executable,
        str(SCRIPT_PATH),
        "--student_csv", student_csv,
        "--teacher_csv", teacher_csv,
        "--group_size", str(group_size_int),
        "--visits_per_group", str(visits_int),
        "--out_dir", out_dir,
    ]

    try:
        status_var.set("Running scheduler...")
        root.update_idletasks()
        subprocess.run(cmd, capture_output=True, text=True, check=True)
        status_var.set("Done.")
        messagebox.showinfo(
            "Success",
            "Scheduling finished successfully.\n\n"
            f"Output folder:\n{out_dir}\n\n"
            "Main output:\n"
            f"{os.path.join(out_dir, 'scheduled_visits.csv')}"
        )
    except subprocess.CalledProcessError as e:
        status_var.set("Failed.")
        details = e.stderr.strip() or e.stdout.strip() or "Unknown error."
        messagebox.showerror("Scheduler failed", details)
    except Exception as e:
        status_var.set("Failed.")
        messagebox.showerror("Unexpected error", str(e))

root = tk.Tk()
root.title("Student Visit Scheduler")
root.geometry("780x380")
root.minsize(780, 380)

padx = 10
pady = 8

student_var = tk.StringVar()
teacher_var = tk.StringVar()
out_var = tk.StringVar(value=str(BASE_DIR / "outputs"))
group_size_var = tk.StringVar(value="2")
visits_var = tk.StringVar(value="2")
status_var = tk.StringVar(value="Ready.")

for i in range(3):
    root.grid_columnconfigure(i, weight=0)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(9, weight=1)

title = tk.Label(
    root,
    text="Student Visit Scheduler",
    font=("Segoe UI", 16, "bold")
)
title.grid(row=0, column=0, columnspan=3, sticky="w", padx=padx, pady=(14, 4))

subtitle = tk.Label(
    root,
    text="Choose the student and teacher CSV files, then click Run Scheduler.",
    font=("Segoe UI", 10)
)
subtitle.grid(row=1, column=0, columnspan=3, sticky="w", padx=padx, pady=(0, 12))

tk.Label(root, text="Student CSV").grid(row=2, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=student_var).grid(row=2, column=1, sticky="ew", padx=padx, pady=pady)
tk.Button(root, text="Browse...", command=lambda: browse_file(student_var, "Select Student CSV")).grid(row=2, column=2, padx=padx, pady=pady)

tk.Label(root, text="Teacher CSV").grid(row=3, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=teacher_var).grid(row=3, column=1, sticky="ew", padx=padx, pady=pady)
tk.Button(root, text="Browse...", command=lambda: browse_file(teacher_var, "Select Teacher CSV")).grid(row=3, column=2, padx=padx, pady=pady)

tk.Label(root, text="Output Folder").grid(row=4, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=out_var).grid(row=4, column=1, sticky="ew", padx=padx, pady=pady)
tk.Button(root, text="Browse...", command=lambda: browse_folder(out_var, "Select Output Folder")).grid(row=4, column=2, padx=padx, pady=pady)

tk.Label(root, text="Group Size").grid(row=5, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=group_size_var, width=10).grid(row=5, column=1, sticky="w", padx=padx, pady=pady)

tk.Label(root, text="Visits per Group").grid(row=6, column=0, sticky="w", padx=padx, pady=pady)
tk.Entry(root, textvariable=visits_var, width=10).grid(row=6, column=1, sticky="w", padx=padx, pady=pady)

button_frame = tk.Frame(root)
button_frame.grid(row=7, column=0, columnspan=3, sticky="w", padx=padx, pady=(16, 8))

tk.Button(
    button_frame,
    text="Run Scheduler",
    command=run_scheduler,
    width=18,
    height=2
).pack(anchor="w")

tk.Label(root, textvariable=status_var).grid(row=8, column=0, columnspan=3, sticky="w", padx=padx, pady=(0, 10))

root.mainloop()
