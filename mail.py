from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.simpledialog import askstring
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Global path to the email file
email_file_path = None

def extract_emails_from_excel(file_path):
    df = pd.read_excel(file_path)
    return df['Parent Email'].tolist()

def upload_email_file():
    global email_file_path
    email_file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if email_file_path:
        emails = extract_emails_from_excel(email_file_path)
        status_bar.config(text=f"✅ Found {len(emails)} emails in the Excel file.")
        option_frame.grid(row=3, column=0, columnspan=2, pady=20)

def upload_overall_attendance():
    if not email_file_path:
        status_bar.config(text="❗Please upload the email Excel file first.")
        return
    overall_file = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if overall_file:
        send_emails_for_attendance(overall_file)

def upload_marks():
    if not email_file_path:
        status_bar.config(text="❗Please upload the email Excel file first.")
        return
    marks_file = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if marks_file:
        send_emails_for_marks(marks_file)

def upload_daily_attendance():
    if not email_file_path:
        status_bar.config(text="❗Please upload the email Excel file first.")
        return
    daily_file = askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if daily_file:
        date = askstring("Enter Date", "Enter the date (e.g. 15-04-2025):")
        if date:
            send_emails_for_daily_attendance(daily_file, date)

def send_email(subject, body, to_email):
    sender_email = "jaykolate529@gmail.com"
    sender_password = "toxkvlwsopwbgowj"
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        server.sendmail(sender_email, to_email, msg.as_string())
        server.quit()
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")

def send_emails_for_attendance(file_path):
    df = pd.read_excel(file_path)
    emails = extract_emails_from_excel(email_file_path)
    for i, row in df.iterrows():
        if row['Attendance (%)'] < 75:
            subject = f"Attendance Alert for {row['Student Name']}"
            body = f"Dear Parent,\n\n{row['Student Name']} has an attendance of {row['Attendance (%)']}%.\nPlease ensure regular attendance.\n\nRegards,\nSchool"
            send_email(subject, body, emails[i])
            print(f"✅ Email sent to {emails[i]}")

def send_emails_for_marks(file_path):
    df = pd.read_excel(file_path)
    emails = extract_emails_from_excel(email_file_path)
    subjects = ['DSA', 'MP', 'M3', 'SE', 'PPL']
    for i, row in df.iterrows():
        low_scores = [f"{subject}: {row[subject]}" for subject in subjects if row[subject] < 40]
        if low_scores:
            all_scores = "\n".join([f"{subject}: {row[subject]}" for subject in subjects])
            subject_line = f"Marks Alert for {row['Student Name']}"
            body = (
                f"Dear Parent,\n\n"
                f"{row['Student Name']} has scored less than 40 in the following subjects:\n" +
                "\n".join(low_scores) +
                "\n\nHere are all the subject scores:\n" +
                all_scores +
                "\n\nPlease take the necessary steps to help improve.\n\nRegards,\nSchool"
            )
            send_email(subject_line, body, emails[i])
            print(f"✅ Email sent to {emails[i]}")

def send_emails_for_daily_attendance(file_path, date):
    df = pd.read_excel(file_path)
    emails = extract_emails_from_excel(email_file_path)
    if date not in df.columns:
        status_bar.config(text="❗Date column not found in file.")
        return
    for i, row in df.iterrows():
        if row[date] == 0:
            subject = f"Absence Alert for {row['Student Name']} on {date}"
            body = f"Dear Parent,\n\n{row['Student Name']} was absent on {date}. Please ensure regular attendance.\n\nRegards,\nSchool"
            send_email(subject, body, emails[i])
            print(f"✅ Email sent to {emails[i]}")

# ---------------- GUI ----------------
root = Tk()
root.title("Student Performance Monitoring System")
root.configure(bg="white")

# Main Frame to center all content
main_frame = Frame(root, bg="white")
main_frame.pack(expand=True)

# Status bar
status_bar = Label(main_frame, text="Ready to upload Email Excel", relief=SUNKEN, anchor=W,
                   bg="#f0f8ff", font=("Arial", 10), width=60)
status_bar.grid(row=0, column=0, columnspan=2, pady=10)

# Upload Email Button
upload_email_button = Button(main_frame, text="Upload Email Excel File", command=upload_email_file,
                             relief="solid", width=30, height=2, bg="#87CEEB", fg="black", font=("Arial", 12))
upload_email_button.grid(row=1, column=0, columnspan=2, pady=10)

# Frame for file options
option_frame = Frame(main_frame, bg="white")

option_label = Label(option_frame, text="Select Excel File Type to Upload", font=("Arial", 12), bg="white")
option_label.grid(row=0, column=0, columnspan=2, pady=10)

overall_button = Button(option_frame, text="Upload Overall Attendance Excle file", command=upload_overall_attendance,
                        relief="solid", width=25, height=2, bg="#87CEEB", fg="black", font=("Arial", 11))
overall_button.grid(row=1, column=0, padx=10, pady=5)

marks_button = Button(option_frame, text="Upload Marks Excel file", command=upload_marks,
                      relief="solid", width=25, height=2, bg="#87CEEB", fg="black", font=("Arial", 11))
marks_button.grid(row=1, column=1, padx=10, pady=5)

daily_button = Button(option_frame, text="Upload Daily Attendance Excel file", command=upload_daily_attendance,
                      relief="solid", width=25, height=2, bg="#87CEEB", fg="black", font=("Arial", 11))
daily_button.grid(row=2, column=0, columnspan=2, pady=10)

# Final window centering
window_width = 600
window_height = 400
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_left = int((screen_width - window_width) / 2)
position_top = int((screen_height - window_height) / 2)
root.geometry(f"{window_width}x{window_height}+{position_left}+{position_top}")

# Start the GUI
root.mainloop()
