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
label_select.pack(side=tk.LEFT)

entry = tk.Entry(frame, width=40)
entry.pack(side=tk.LEFT, padx=5)

browse_button = tk.Button(frame, text="Browse", command=select_file)
browse_button.pack(side=tk.LEFT, padx=5)

button_frame = tk.Frame(root)
button_frame.pack(pady=5)

upload_button = tk.Button(button_frame, text="Upload", command=upload_file)
upload_button.pack(side=tk.LEFT, padx=5)

upload_and_book_button = tk.Button(button_frame, text="Upload and Book Event", command=upload_file)
upload_and_book_button.pack(side=tk.LEFT, padx=5)

exit_button = tk.Button(button_frame, text="Exit", command=exit_program)
exit_button.pack(side=tk.LEFT, padx=5)

root.mainloop()
