import tkinter as tk
from tkinter import messagebox, simpledialog, ttk, filedialog
import json
import os
from datetime import datetime, timedelta
import pandas as pd
import subprocess  # For opening file location
import smtplib  # For sending emails
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- File Configurations ---
CONFIG_FILE = "config.json"
ADMINS_FILE = "admins.json"
EMAIL_CONFIG_FILE = "email_config.json"  # New file for email settings
RECIPIENTS_FILE = "recipients.json"  # New file for notification recipients

# Construct the path to save the Excel file in the user's 'Documents' folder
documents_path = os.path.join(os.path.expanduser("~"), "Documents")
RECORDS_EXCEL_FILE = os.path.join(documents_path, "student_attendance_records.xlsx")

STUDENTS_FILE = "students.json"  # Students file remains in the script's directory

# --- Developer Credentials ---
DEV_USERNAME = "ADMIN"
DEV_PASSWORD = "BRAIN FEMI"

# --- App Info ---
APP_TITLE = "Attendance and Daily Event Book"
CREATOR_NOTE = "App Created by Charles Oluwafemi Ademokun"

# --- Global Variables for Alerts ---
LAST_ALERTED_MORNING = None
LAST_ALERTED_AFTERNOON = None
# Define daily attendance times (e.g., 9 AM for Morning, 1 PM for Afternoon)
ATTENDANCE_SCHEDULE = {
    "Morning": "09:00",
    "Afternoon": "13:00"
}


# --- Initialize App ---
def initialize_app():
    """
    Initializes the application by creating necessary configuration and data files.
    Ensures the Excel attendance file exists in the 'Documents' folder with headers.
    """
    print(f"DEBUG: Initializing app. Records Excel file path: {RECORDS_EXCEL_FILE}")

    if not os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "w") as f:
            json.dump({"first_use": datetime.now().strftime("%Y-%m-%d")}, f)

    os.makedirs(documents_path, exist_ok=True)

    # Ensure the Excel file exists with "Session" header
    if not os.path.exists(RECORDS_EXCEL_FILE):
        df = pd.DataFrame(columns=["Date", "Session", "Student", "Status"])  # Added 'Session' column
        df.to_excel(RECORDS_EXCEL_FILE, index=False)
        messagebox.showinfo("File Creation", f"Attendance Excel file created at: {RECORDS_EXCEL_FILE}")
    else:
        # Check if 'Session' column exists, add if not (for existing files)
        try:
            df = pd.read_excel(RECORDS_EXCEL_FILE)
            if "Session" not in df.columns:
                df.insert(1, "Session", "Morning")  # Insert 'Session' after 'Date', default to 'Morning'
                df.to_excel(RECORDS_EXCEL_FILE, index=False)
                messagebox.showinfo("File Update", "Added 'Session' column to existing attendance records.")
        except Exception as e:
            messagebox.showwarning("File Read Error",
                                   f"Could not read existing Excel file to check for 'Session' column: {e}")
        print(f"DEBUG: Attendance Excel file already exists at: {RECORDS_EXCEL_FILE}")

    # Initialize email config and recipients files if they don't exist
    if not os.path.exists(EMAIL_CONFIG_FILE):
        with open(EMAIL_CONFIG_FILE, "w") as f:
            json.dump({"smtp_server": "", "smtp_port": 587, "sender_email": "", "sender_password": ""}, f)
    if not os.path.exists(RECIPIENTS_FILE):
        with open(RECIPIENTS_FILE, "w") as f:
            json.dump([], f)  # Store recipients as a list of emails


# --- Expiry Check ---
def is_expired():
    """Checks if the application's one-year usage period has expired."""
    with open(CONFIG_FILE) as f:
        data = json.load(f)
        first_use = datetime.strptime(data["first_use"], "%Y-%m-%d")
        return datetime.now() > first_use + timedelta(days=365)


def reset_expiry():
    """Resets the application's expiry date (developer function)."""
    with open(CONFIG_FILE, "w") as f:
        json.dump({"first_use": datetime.now().strftime("%Y-%m-%d")}, f)
    messagebox.showinfo("Reset", "Expiry date has been reset by the developer.")


# --- Admin Handlers ---
def load_admins():
    """Loads admin user credentials from the admins.json file."""
    if not os.path.exists(ADMINS_FILE):
        with open(ADMINS_FILE, "w") as f:
            json.dump({}, f)
    with open(ADMINS_FILE) as f:
        return json.load(f)


def save_admins(admins):
    """Saves admin user credentials to the admins.json file."""
    with open(ADMINS_FILE, "w") as f:
        json.dump(admins, f)


# --- Email Configuration Handlers ---
def load_email_config():
    """Loads email configuration from the email_config.json file."""
    if not os.path.exists(EMAIL_CONFIG_FILE):
        # Default empty config
        return {"smtp_server": "", "smtp_port": 587, "sender_email": "", "sender_password": ""}
    with open(EMAIL_CONFIG_FILE, "r") as f:
        return json.load(f)


def save_email_config(config):
    """Saves email configuration to the email_config.json file."""
    with open(EMAIL_CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=4)


def load_recipients():
    """Loads notification recipients from the recipients.json file."""
    if not os.path.exists(RECIPIENTS_FILE):
        return []
    with open(RECIPIENTS_FILE, "r") as f:
        return json.load(f)


def save_recipients(recipients):
    """Saves notification recipients to the recipients.json file."""
    with open(RECIPIENTS_FILE, "w") as f:
        json.dump(recipients, f, indent=4)


def configure_email_settings():
    """Allows setting up SMTP server, sender email, and password."""
    config = load_email_config()

    win = tk.Toplevel()
    win.title("Email Settings")
    win.geometry("450x350")
    win.configure(bg="#e6f2ff")

    tk.Label(win, text="SMTP Server:", bg="#e6f2ff").pack(pady=5)
    smtp_server_entry = tk.Entry(win, width=40)
    smtp_server_entry.insert(0, config.get("smtp_server", ""))
    smtp_server_entry.pack()

    tk.Label(win, text="SMTP Port (e.g., 587 for TLS, 465 for SSL):", bg="#e6f2ff").pack(pady=5)
    smtp_port_entry = tk.Entry(win, width=40)
    smtp_port_entry.insert(0, str(config.get("smtp_port", 587)))
    smtp_port_entry.pack()

    tk.Label(win, text="Sender Email:", bg="#e6f2ff").pack(pady=5)
    sender_email_entry = tk.Entry(win, width=40)
    sender_email_entry.insert(0, config.get("sender_email", ""))
    sender_email_entry.pack()

    tk.Label(win, text="Sender Password (App Password Recommended):", bg="#e6f2ff").pack(pady=5)
    sender_password_entry = tk.Entry(win, width=40, show="*")
    sender_password_entry.insert(0, config.get("sender_password", ""))
    sender_password_entry.pack()

    def save_settings():
        new_config = {
            "smtp_server": smtp_server_entry.get().strip(),
            "smtp_port": int(smtp_port_entry.get().strip()),
            "sender_email": sender_email_entry.get().strip(),
            "sender_password": sender_password_entry.get().strip()
        }
        save_email_config(new_config)
        messagebox.showinfo("Saved", "Email settings updated. Remember to use an App Password if using Gmail/Outlook.")
        win.destroy()

    tk.Button(win, text="Save Settings", command=save_settings, bg="#4CAF50", fg="white").pack(pady=15)
    tk.Label(win, text="For Gmail, you might need to generate an 'App password' if 2FA is enabled.", bg="#e6f2ff",
             font=("Arial", 8)).pack()
    tk.Label(win, text="Search 'Gmail App Passwords' for instructions.", bg="#e6f2ff", font=("Arial", 8)).pack()


def manage_recipients():
    """Allows managing email recipients for absence notifications."""
    recipients = load_recipients()

    win = tk.Toplevel()
    win.title("Manage Email Recipients")
    win.geometry("400x400")
    win.configure(bg="#e6f2ff")

    tk.Label(win, text="Enter recipient emails (one per line):", font=("Arial", 12), bg="#e6f2ff").pack(pady=10)
    text_box = tk.Text(win, width=40, height=15)
    text_box.pack(pady=10)
    text_box.insert(tk.END, "\n".join(recipients))

    def save_recipients_list():
        new_recipients = [email.strip() for email in text_box.get("1.0", tk.END).split("\n") if email.strip()]
        save_recipients(new_recipients)
        messagebox.showinfo("Saved", "Recipient list updated.")
        win.destroy()

    tk.Button(win, text="Save Recipients", command=save_recipients_list, bg="#4CAF50", fg="white").pack(pady=5)
    tk.Label(win, text="Notifications via WhatsApp/Phone Call are not supported.", bg="#e6f2ff", fg="red").pack(pady=5)


def send_email_notification(absent_student_name, session_name, date_str):
    """
    Sends an email notification to configured recipients about an absent student.
    """
    config = load_email_config()
    recipients = load_recipients()

    smtp_server = config.get("smtp_server")
    smtp_port = config.get("smtp_port")
    sender_email = config.get("sender_email")
    sender_password = config.get("sender_password")

    if not all([smtp_server, smtp_port, sender_email, sender_password]) or not recipients:
        print("DEBUG: Email settings or recipients not fully configured. Skipping email notification.")
        return False  # Indicate that email was not sent due to missing config

    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = f"Urgent: Student Absent - {absent_student_name} ({session_name} Session)"

    body = f"Dear Authority,\n\nThis is an automated notification.\n\nStudent: {absent_student_name} is marked ABSENT for the {session_name} session today, {date_str}.\n\nPlease take necessary action.\n\nSincerely,\nAttendance System"
    msg.attach(MIMEText(body, 'plain'))

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()  # Start TLS encryption
            server.login(sender_email, sender_password)
            server.send_message(msg)
        print(f"DEBUG: Email notification sent for {absent_student_name}.")
        return True  # Indicate success
    except Exception as e:
        print(f"ERROR: Failed to send email notification for {absent_student_name}: {e}")
        messagebox.showerror("Email Error",
                             f"Failed to send email for {absent_student_name}. Check email settings. Error: {e}")
        return False  # Indicate failure


def save_attendance(date_entry_widget, session_combobox_widget, entries_dict, window_instance):
    """
    Saves the marked attendance records to the Excel file.
    Updates existing entries for the same date, student, and session, and adds new ones.
    Confirms submission to Excel and sends email for absent students in Morning session.
    """
    print(f"DEBUG: save_attendance function called.")
    date_str = date_entry_widget.get()
    session_name = session_combobox_widget.get()
    current_date = datetime.now().strftime("%Y-%m-%d")

    try:
        datetime.strptime(date_str, "%Y-%m-%d")
    except ValueError:
        messagebox.showerror("Invalid Date", "Please enter the date in Jamboree-MM-DD format.")
        return

    attendance_records = []
    absent_students_morning = []  # List to track absent students for morning notification

    for student, cb in entries_dict.items():
        status = cb.get()
        attendance_records.append({"Date": date_str, "Session": session_name, "Student": student, "Status": status})
        if session_name == "Morning" and status == "Absent":
            absent_students_morning.append(student)

    new_df = pd.DataFrame(attendance_records)
    print(f"DEBUG: New attendance data collected for {session_name} session:\n{new_df}")

    try:
        print(f"DEBUG: Attempting to read existing Excel file from: {RECORDS_EXCEL_FILE}")
        existing_df = pd.read_excel(RECORDS_EXCEL_FILE)
        print("DEBUG: Existing Excel file read successfully.")
    except FileNotFoundError:
        print("DEBUG: Existing Excel file not found. Creating empty DataFrame.")
        existing_df = pd.DataFrame(columns=["Date", "Session", "Student", "Status"])
    except Exception as e:
        print(f"ERROR: Could not read the attendance file: {e}\nAttempting to create a new one.")
        messagebox.showerror("Error Reading Excel",
                             f"Could not read the attendance file: {e}\nAttempting to create a new one.")
        existing_df = pd.DataFrame(columns=["Date", "Session", "Student", "Status"])

    # Combine existing and new data.
    # drop_duplicates with 'keep='last'' ensures that if an entry for the same
    # student on the same date and session exists, the new entry replaces the old one.
    combined_df = pd.concat([existing_df, new_df], ignore_index=True).drop_duplicates(
        subset=['Date', 'Session', 'Student'], keep='last')
    print(f"DEBUG: Combined DataFrame to be saved:\n{combined_df}")

    try:
        print(f"DEBUG: Attempting to save combined DataFrame to Excel at: {RECORDS_EXCEL_FILE}")
        combined_df.to_excel(RECORDS_EXCEL_FILE, index=False)
        print("DEBUG: Excel file saved successfully.")
        messagebox.showinfo("Attendance Saved",
                            f"Today's attendance ({date_str}, {session_name} Session) has been successfully submitted and saved to the Excel file.")
        window_instance.destroy()  # Close the attendance window
        open_file_location(RECORDS_EXCEL_FILE)  # Open file location after saving

        # --- Send email notification for absent students in Morning session ---
        global LAST_ALERTED_MORNING
        if session_name == "Morning" and date_str == current_date:
            # Check if morning attendance was just saved and notifications haven't been sent yet for today
            # This logic might need refinement if 'LAST_ALERTED_MORNING' is meant for teacher alerts
            # rather than confirming email has been sent for a specific session.
            # For simplicity, if attendance is saved for morning, send notifications for absentees.
            # A more robust check might involve tracking sent emails per student per session.
            for student in absent_students_morning:
                send_email_notification(student, session_name, date_str)  # Pass session and date for email content

    except Exception as e:
        print(f"ERROR: Could not save attendance to Excel file: {e}")
        messagebox.showerror("Error Saving Excel", f"Could not save attendance to Excel file: {e}")


# --- Login Function ---
def login():
    """Handles user login authentication."""
    if is_expired():
        messagebox.showerror("Expired", "This app has expired after 1 year of use. Please contact the developer.")
        return

    user = username_entry.get()
    pwd = password_entry.get()

    if user == DEV_USERNAME and pwd == DEV_PASSWORD:
        messagebox.showinfo("Developer Login", "Developer logged in successfully.")
        root.destroy()
        launch_dashboard("developer")
        return

    admins = load_admins()
    if user in admins and admins[user] == pwd:
        messagebox.showinfo("Admin Login", f"Welcome Admin: {user}")
        root.destroy()
        launch_dashboard("admin", current_user=user)
    else:
        messagebox.showerror("Login Failed", "Invalid username or password.")


# --- Admin Management ---
def manage_admins(current_role, current_user=None):
    """
    Allows developers/admins to add, remove, or change passwords for admin users.
    'admin' role can only change their own password.
    """
    admins = load_admins()

    def refresh_list():
        """Refreshes the listbox displaying current admin users."""
        admin_list.delete(0, tk.END)
        for a in admins:
            admin_list.insert(tk.END, a)

    def add_admin():
        """Prompts for new admin username and password and adds them."""
        new_user = simpledialog.askstring("New Admin", "Enter username:")
        new_pass = simpledialog.askstring("New Admin", "Enter password:")
        if new_user and new_pass:
            admins[new_user] = new_pass
            save_admins(admins)
            refresh_list()
            messagebox.showinfo("Success", f"Admin '{new_user}' added.")

    def remove_admin():
        """Removes the selected admin user."""
        selected = admin_list.curselection()
        if selected:
            user_to_remove = admin_list.get(selected)
            if current_role == "admin" and user_to_remove == current_user:
                messagebox.showerror("Error", "You cannot remove yourself.")
                return
            confirm = messagebox.askyesno("Confirm Removal", f"Are you sure you want to remove '{user_to_remove}'?")
            if confirm:
                del admins[user_to_remove]
                save_admins(admins)
                refresh_list()
                messagebox.showinfo("Success", f"Admin '{user_to_remove}' removed.")
        else:
            messagebox.showerror("Selection Error", "Please select an admin to remove.")

    def change_password():
        """Allows the current user to change their password."""
        if current_user:
            new_pass = simpledialog.askstring("Change Password", "Enter new password:")
            if new_pass:
                admins[current_user] = new_pass
                save_admins(admins)
                messagebox.showinfo("Success", "Password updated.")
        else:
            messagebox.showerror("Error", "No current user detected for password change.")

    win = tk.Toplevel()
    win.title("Manage Admins")
    win.geometry("400x300")
    win.configure(bg="#e6f2ff")

    tk.Label(win, text="Admin Users:", font=("Arial", 12), bg="#e6f2ff").pack(pady=5)
    admin_list = tk.Listbox(win, width=40, height=10)
    admin_list.pack(pady=5)

    button_frame = tk.Frame(win, bg="#e6f2ff")
    button_frame.pack(pady=5)

    tk.Button(button_frame, text="Add Admin", command=add_admin, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Remove Selected", command=remove_admin, bg="#F44336", fg="white").pack(side=tk.LEFT,
                                                                                                         padx=5)

    if current_role == "admin":
        tk.Button(button_frame, text="Change My Password", command=change_password, bg="#2196F3", fg="white").pack(
            side=tk.LEFT, padx=5)

    refresh_list()


# --- Advanced Settings ---
def advanced_settings():
    """Provides advanced settings options (e.g., reset expiry for developers)."""
    win = tk.Toplevel()
    win.title("Advanced Settings")
    win.geometry("300x150")
    win.configure(bg="#e6f2ff")
    tk.Button(win, text="Reset Expiry Date", command=reset_expiry, bg="#FFC107", fg="black").pack(pady=20)


# --- Student Registration ---
def register_students():
    """Allows registration of student names, one per line."""

   
