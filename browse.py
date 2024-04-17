import tkinter as tk
from tkinter import filedialog
import pandas as pd
import win32com.client as win32

def read_email_ids_from_excel(file_path):
    df = pd.read_excel(file_path)
    email_ids = df['Email'].tolist()
    return email_ids

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def upload_file():
    file_path = entry.get()  # Get the file path from the entry box
    if file_path:
        email_ids = read_email_ids_from_excel(file_path)
        subject = 'Meeting'
        start_time = '2024-04-20 09:00'
        end_time = '2024-04-20 10:00'
        book_calendar_event(email_ids, subject, start_time, end_time)
        print("Calendar event booked successfully.")
    else:
        print("No file selected")

def exit_program():
    root.destroy()

def book_calendar_event(email_ids, subject, start_time, end_time):
    outlook = win32.Dispatch('Outlook.Application')
    appointment = outlook.CreateItem(1)  # 1 represents an appointment item

    appointment.Subject = subject
    appointment.Start = start_time
    appointment.End = end_time

    for email_id in email_ids:
        appointment.Recipients.Add(email_id)

    appointment.Save()
    appointment.Send()

root = tk.Tk()
root.title("Upload Excel File and Book Calendar Event")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

label_select = tk.Label(frame, text="Select Excel File:")
label_select.pack(side=tk.TOP, anchor='w')

entry = tk.Entry(frame, width=40)
entry.pack(side=tk.LEFT, padx=5)

browse_button = tk.Button(frame, text="Browse", command=select_file)
browse_button.pack(side=tk.LEFT)

upload_button = tk.Button(root, text="Upload and Book Event", command=upload_file)
upload_button.pack(side=tk.TOP, padx=5, pady=5)

exit_button = tk.Button(root, text="Exit", command=exit_program)
exit_button.pack(side=tk.TOP, padx=5, pady=5)

root.mainloop()


from datetime import datetime, timedelta

# Get today's date
today_date = datetime.now().date()

# Define the time to add (9:00 AM)
time_to_add = timedelta(hours=9)

# Combine today's date with the time to add
new_datetime = datetime.combine(today_date, datetime.min.time()) + time_to_add

print("New datetime (9:00 AM):", new_datetime)
