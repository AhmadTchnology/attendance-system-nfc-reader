import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sqlite3
from datetime import datetime
from openpyxl import Workbook, load_workbook  # For Excel file handling
from PIL import Image, ImageTk  # For adding images
import os  # To scan for .db files
import sys

# Helper function to handle file paths for PyInstaller
def resource_path(relative_path):
    """Get the absolute path to a resource, works for dev and PyInstaller."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Global variables
current_db = None
conn = None
first_major = None  # To store the major of the first student who taps their tag
first_stage = None  # To store the stage of the first student
first_study = None  # To store the study of the first student
first_group = None  # To store the group of the first student

def create_db_connection(db_file):
    """Create a connection to the SQLite database."""
    global conn
    try:
        conn = sqlite3.connect(db_file)
        return True
    except sqlite3.Error as e:
        messagebox.showerror("Database Error", str(e))
        return False

def initialize_db():
    """Initialize the database with the required tables."""
    if conn:
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS students (
                student_id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                major TEXT NOT NULL,
                stage TEXT NOT NULL,
                study TEXT NOT NULL,
                group_name TEXT NOT NULL
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id TEXT,
                name TEXT,
                major TEXT,
                stage TEXT,
                study TEXT,
                group_name TEXT,
                timestamp TEXT,
                attended INTEGER,
                FOREIGN KEY(student_id) REFERENCES students(student_id)
            )
        """)
        conn.commit()

def load_db_from_dropdown(event=None):
    """Load a database selected from the dropdown."""
    global current_db, first_major, first_stage, first_study, first_group
    selected_db = db_dropdown.get().strip()
    if not selected_db:
        messagebox.showwarning("Input Error", "No database selected.")
        return
    db_file = os.path.join("Students Databases", selected_db)
    if create_db_connection(db_file):
        current_db = db_file
        initialize_db()
        first_major = None  # Reset filters when loading a new database
        first_stage = None
        first_study = None
        first_group = None
        messagebox.showinfo("Success", f"Database loaded: {db_file}")
    else:
        messagebox.showerror("Error", "Failed to load database.")

def populate_db_dropdown():
    """Populate the dropdown with .db files from the 'Students Databases' folder."""
    db_files = []
    try:
        if not os.path.exists("Students Databases"):
            os.makedirs("Students Databases")  # Create the folder if it doesn't exist
        db_files = [f for f in os.listdir("Students Databases") if f.endswith(".db")]
    except Exception as e:
        print(f"Error reading 'Students Databases' folder: {e}")
    db_dropdown['values'] = db_files
    if db_files:
        db_dropdown.current(0)  # Select the first database by default
    else:
        messagebox.showwarning("Warning", "No .db files found in 'Students Databases' folder.")

def create_db():
    """Create a new .db file."""
    global current_db, first_major, first_stage, first_study, first_group
    db_file = filedialog.asksaveasfilename(
        title="Create New Database File",
        defaultextension=".db",
        filetypes=[("SQLite Database", "*.db")]
    )
    if db_file:
        if create_db_connection(db_file):
            current_db = db_file
            initialize_db()
            first_major = None  # Reset filters when creating a new database
            first_stage = None
            first_study = None
            first_group = None
            messagebox.showinfo("Success", f"New database created: {db_file}")
            # Move the new database to the 'Students Databases' folder
            target_folder = "Students Databases"
            if not os.path.exists(target_folder):
                os.makedirs(target_folder)
            target_path = os.path.join(target_folder, os.path.basename(db_file))
            os.replace(db_file, target_path)
            populate_db_dropdown()  # Refresh the dropdown
        else:
            messagebox.showerror("Error", "Failed to create database.")

def add_student():
    """Add a new student to the database."""
    if not conn:
        messagebox.showerror("Error", "No database loaded.")
        return
    student_id = entry_id.get().strip()
    name = entry_name.get().strip()
    major = entry_major.get().strip()
    stage = entry_stage.get().strip()
    study = entry_study.get().strip()
    group = entry_group.get().strip()
    if not student_id or not name or not major or not stage or not study or not group:
        messagebox.showwarning("Input Error", "All fields are required.")
        return
    try:
        cursor = conn.cursor()
        cursor.execute("""
            INSERT INTO students (student_id, name, major, stage, study, group_name)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (student_id, name, major, stage, study, group))
        conn.commit()
        messagebox.showinfo("Success", "Student added successfully.")
        clear_entries()
    except sqlite3.IntegrityError:
        messagebox.showerror("Error", "Student ID already exists.")

def clear_entries():
    """Clear input fields."""
    entry_id.delete(0, tk.END)
    entry_name.delete(0, tk.END)
    entry_major.delete(0, tk.END)
    entry_stage.delete(0, tk.END)
    entry_study.delete(0, tk.END)
    entry_group.delete(0, tk.END)

def import_students_from_excel():
    """Import students from an Excel file into the database."""
    if not conn:
        messagebox.showerror("Error", "No database loaded.")
        return
    excel_file = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not excel_file:
        return  # Exit if no file is selected
    try:
        wb = load_workbook(excel_file)
        ws = wb.active
        cursor = conn.cursor()
        success_count = 0
        duplicate_count = 0
        for row in ws.iter_rows(min_row=2, values_only=True):  # Skip header row
            student_id, name, major, stage, study, group = row[0], row[1], row[2], row[3], row[4], row[5]
            if not student_id or not name or not major or not stage or not study or not group:
                continue  # Skip empty rows
            try:
                cursor.execute("""
                    INSERT INTO students (student_id, name, major, stage, study, group_name)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (student_id, name, major, stage, study, group))
                success_count += 1
            except sqlite3.IntegrityError:
                duplicate_count += 1  # Count duplicates but don't stop
        conn.commit()
        message = f"Successfully imported {success_count} students.\n"
        if duplicate_count > 0:
            message += f"{duplicate_count} duplicate entries were skipped."
        messagebox.showinfo("Import Complete", message)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to import students: {str(e)}")

def record_attendance(event=None):
    """Record attendance using NFC and update the dashboard."""
    global first_major, first_stage, first_study, first_group
    if not conn:
        return  # Silently exit if no database is loaded
    nfc_id = entry_nfc.get().strip()
    if not nfc_id:
        return  # Silently exit if no NFC ID is detected
    entry_nfc.delete(0, tk.END)
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM students WHERE student_id=?", (nfc_id,))
    student = cursor.fetchone()
    if not student:
        return  # Silently exit if the student is not found

    # Check if the student's major, stage, study, and group match the filters
    if first_major is None:
        first_major = student[2]  # Set the first_major to the major of the first student
        first_stage = student[3]
        first_study = student[4]
        first_group = student[5]
        messagebox.showinfo("Info", f"Filters set to: Major={first_major}, Stage={first_stage}, Study={first_study}, Group={first_group}")
    elif student[2] != first_major or student[3] != first_stage or student[4] != first_study or student[5] != first_group:
        messagebox.showwarning("Filter Mismatch", f"Attendance is currently restricted to students with:\nMajor={first_major}, Stage={first_stage}, Study={first_study}, Group={first_group}.")
        return

    # Check if the student has already been marked as attended today
    today = datetime.now().strftime("%Y-%m-%d")
    cursor.execute("""
        SELECT * FROM attendance 
        WHERE student_id=? AND DATE(timestamp)=?
    """, (nfc_id, today))
    existing_record = cursor.fetchone()
    if existing_record:
        messagebox.showwarning("Already Attended", f"{student[1]} has already been marked as attended today.")
        return

    # Insert the new attendance record
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cursor.execute("""
        INSERT INTO attendance (student_id, name, major, stage, study, group_name, timestamp, attended)
        VALUES (?, ?, ?, ?, ?, ?, ?, 1)
    """, (student[0], student[1], student[2], student[3], student[4], student[5], timestamp))
    conn.commit()

    # Update the dashboard
    dashboard_tree.insert("", tk.END, values=(student[1], student[2], student[3], student[4], student[5], timestamp, "Yes"))

def export_attendance():
    """Export attendance data to an Excel file, filtered by the first student's major, stage, study, and group."""
    global first_major, first_stage, first_study, first_group
    if not conn:
        messagebox.showerror("Error", "No database loaded.")
        return
    if first_major is None:
        messagebox.showwarning("Warning", "No attendance recorded yet. Cannot determine the filters.")
        return
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.name, s.major, s.stage, s.study, s.group_name, a.timestamp, CASE WHEN a.attended = 1 THEN 'Yes' ELSE 'No' END AS attended
        FROM students s
        LEFT JOIN attendance a ON s.student_id = a.student_id
        WHERE s.major = ? AND s.stage = ? AND s.study = ? AND s.group_name = ?
    """, (first_major, first_stage, first_study, first_group))
    rows = cursor.fetchall()
    if not rows:
        messagebox.showwarning("Warning", f"No attendance data found for filters: Major={first_major}, Stage={first_stage}, Study={first_study}, Group={first_group}.")
        return
    wb = Workbook()
    ws = wb.active
    ws.title = "Attendance"
    ws.append(["Name", "Major", "Stage", "Study", "Group", "Timestamp", "Attended"])
    for row in rows:
        ws.append(row)
    export_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        initialfile=f"{datetime.now().strftime('%Y-%m-%d')}_{first_major}_{first_stage}_{first_study}_{first_group}"
    )
    if export_file:
        wb.save(export_file)
        messagebox.showinfo("Success", f"Attendance exported to {export_file}.")

def reset_attendance():
    """Reset all attendance data and clear the filters."""
    global first_major, first_stage, first_study, first_group
    if not conn:
        messagebox.showerror("Error", "No database loaded.")
        return
    # Clear all attendance records
    cursor = conn.cursor()
    cursor.execute("DELETE FROM attendance")
    conn.commit()
    # Clear the dashboard
    dashboard_tree.delete(*dashboard_tree.get_children())
    # Reset the filters
    first_major = None
    first_stage = None
    first_study = None
    first_group = None
    messagebox.showinfo("Success", "Attendance data and filters have been reset.")

# GUI Setup
root = tk.Tk()
root.title("📊 Real-Time Attendance Dashboard")
root.geometry("1200x1000")
root.resizable(False, False)

# Styling
style = ttk.Style()
style.configure("TLabel", font=("Segoe UI", 12), padding=5)
style.configure("TButton", font=("Segoe UI", 12), padding=8, relief="flat")
style.map("TButton", background=[("active", "#219150"), ("pressed", "#1e5dbd")])
style.configure("Header.TLabel", font=("Segoe UI", 20, "bold"), foreground="#fff", background="#2575fc", padding=10)
style.configure("Dashboard.TLabel", font=("Segoe UI", 14, "bold"), foreground="#333")

# Header
header_frame = tk.Frame(root, bg="#2575fc", height=60)
header_frame.pack(fill="x")
ttk.Label(header_frame, text="📊 Real-Time Attendance Dashboard", style="Header.TLabel").pack()

# Main Frame
main_frame = ttk.Frame(root, padding=20)
main_frame.pack(fill="both", expand=True)

# Left Panel (Logos)
left_panel = ttk.Frame(main_frame)
left_panel.grid(row=0, column=0, sticky="ns", padx=10, pady=10)

# Add Any Logo
try:
    university_logo = Image.open(resource_path("Any_Image_You_Want.png")).resize((260, 300))
    university_logo_img = ImageTk.PhotoImage(university_logo)
    university_label = ttk.Label(left_panel, image=university_logo_img)
    university_label.image = university_logo_img
    university_label.pack(pady=10)
except Exception as e:
    print("University logo not found:", e)
    ttk.Label(left_panel, text="Any_Logo", font=("Segoe UI", 14)).pack(pady=10)

# Add Any Logo
try:
    engineering_logo = Image.open(resource_path("Any_Image_You_Want.png")).resize((270, 260))
    engineering_logo_img = ImageTk.PhotoImage(engineering_logo)
    engineering_label = ttk.Label(left_panel, image=engineering_logo_img)
    engineering_label.image = engineering_logo_img
    engineering_label.pack(pady=10)
except Exception as e:
    print("Engineering logo not found:", e)
    ttk.Label(left_panel, text="Any_Logo", font=("Segoe UI", 14)).pack(pady=10)

# Add "Made By Ahmad Tchnology" Label
made_by_label = ttk.Label(left_panel, text="Made By Ahmad Tchnology", font=("Segoe UI", 12, "italic"), foreground="#555")
made_by_label.pack(side=tk.BOTTOM, pady=10)

# Right Panel (Main Functionality)
right_panel = ttk.Frame(main_frame)
right_panel.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

# Database Selector
database_frame = ttk.Frame(right_panel)
database_frame.grid(row=0, column=0, columnspan=2, pady=10)
ttk.Label(database_frame, text="Select Database:").pack(side=tk.LEFT, padx=5)

# Dropdown for Database Selection
db_dropdown = ttk.Combobox(database_frame, state="readonly", width=30)
db_dropdown.pack(side=tk.LEFT, padx=5)
populate_db_dropdown()  # Populate the dropdown with .db files
db_dropdown.bind("<<ComboboxSelected>>", load_db_from_dropdown)  # Load the selected database

# Create New Database Button
btn_create_db = ttk.Button(database_frame, text="Create New Database", command=create_db)
btn_create_db.pack(side=tk.LEFT, padx=5)

# Student Entry Fields
entry_frame = ttk.Frame(right_panel)
entry_frame.grid(row=1, column=0, columnspan=2, pady=10)
ttk.Label(entry_frame, text="Student ID:").grid(row=0, column=0, sticky=tk.W, pady=5)
entry_id = ttk.Entry(entry_frame, width=30)
entry_id.grid(row=0, column=1, pady=5, sticky=tk.W)
ttk.Label(entry_frame, text="Name:").grid(row=1, column=0, sticky=tk.W, pady=5)
entry_name = ttk.Entry(entry_frame, width=30)
entry_name.grid(row=1, column=1, pady=5, sticky=tk.W)
ttk.Label(entry_frame, text="Major:").grid(row=2, column=0, sticky=tk.W, pady=5)
entry_major = ttk.Entry(entry_frame, width=30)
entry_major.grid(row=2, column=1, pady=5, sticky=tk.W)
ttk.Label(entry_frame, text="Stage:").grid(row=3, column=0, sticky=tk.W, pady=5)
entry_stage = ttk.Entry(entry_frame, width=30)
entry_stage.grid(row=3, column=1, pady=5, sticky=tk.W)
ttk.Label(entry_frame, text="Study:").grid(row=4, column=0, sticky=tk.W, pady=5)
entry_study = ttk.Entry(entry_frame, width=30)
entry_study.grid(row=4, column=1, pady=5, sticky=tk.W)
ttk.Label(entry_frame, text="Group:").grid(row=5, column=0, sticky=tk.W, pady=5)
entry_group = ttk.Entry(entry_frame, width=30)
entry_group.grid(row=5, column=1, pady=5, sticky=tk.W)
btn_add_student = ttk.Button(entry_frame, text="Add Student", command=add_student)
btn_add_student.grid(row=6, column=0, columnspan=2, pady=10)

# Import Students Button
btn_import_students = ttk.Button(entry_frame, text="Import Students from Excel", command=import_students_from_excel)
btn_import_students.grid(row=7, column=0, columnspan=2, pady=10)

# NFC Entry Field
nfc_frame = ttk.Frame(right_panel)
nfc_frame.grid(row=2, column=0, columnspan=2, pady=10)
ttk.Label(nfc_frame, text="Scan NFC Card:").grid(row=0, column=0, sticky=tk.W, pady=5)
entry_nfc = ttk.Entry(nfc_frame, width=30)
entry_nfc.grid(row=0, column=1, pady=5, sticky=tk.W)
entry_nfc.bind("<Return>", record_attendance)

# Buttons Row (Record, Export, Reset)
buttons_frame = ttk.Frame(nfc_frame)
buttons_frame.grid(row=1, column=0, columnspan=2, pady=10)
btn_record_attendance = ttk.Button(buttons_frame, text="Record Attendance", command=record_attendance)
btn_record_attendance.pack(side=tk.LEFT, padx=5)
btn_export_attendance = ttk.Button(buttons_frame, text="📤 Export Attendance", command=export_attendance, style="Green.TButton")
btn_export_attendance.pack(side=tk.LEFT, padx=5)
btn_reset_attendance = ttk.Button(buttons_frame, text="🔄 Reset Attendance", command=reset_attendance, style="Red.TButton")
btn_reset_attendance.pack(side=tk.LEFT, padx=5)

# Button Styles
style.configure("Green.TButton", background="#27ae60")
style.configure("Red.TButton", background="#e74c3c")

# Dashboard (Treeview to display student names and timestamps)
dashboard_frame = ttk.Frame(right_panel)
dashboard_frame.grid(row=3, column=0, columnspan=2, pady=10)
ttk.Label(dashboard_frame, text="Attendance Dashboard", style="Dashboard.TLabel").pack()
dashboard_tree = ttk.Treeview(dashboard_frame, columns=("Name", "Major", "Stage", "Study", "Group", "Timestamp", "Attended"), show="headings", height=10)
dashboard_tree.heading("Name", text="Name")
dashboard_tree.heading("Major", text="Major")
dashboard_tree.heading("Stage", text="Stage")
dashboard_tree.heading("Study", text="Study")
dashboard_tree.heading("Group", text="Group")
dashboard_tree.heading("Timestamp", text="Timestamp")
dashboard_tree.heading("Attended", text="Attended")
dashboard_tree.column("Name", width=150)
dashboard_tree.column("Major", width=100)
dashboard_tree.column("Stage", width=100)
dashboard_tree.column("Study", width=100)
dashboard_tree.column("Group", width=100)
dashboard_tree.column("Timestamp", width=150)
dashboard_tree.column("Attended", width=100)
dashboard_tree.pack()

# Run the application
root.mainloop()