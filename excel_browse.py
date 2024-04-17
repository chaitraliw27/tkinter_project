import tkinter as tk
from tkinter import filedialog
import pandas as pd

def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    if file_path:
        entry.delete(0, tk.END)
        entry.insert(0, file_path)

def upload_file():
    file_path = entry.get()  # Get the file path from the entry box
    if file_path:
        # Add your file upload logic here
        print("File uploaded:", file_path)
    else:
        print("No file selected")

def exit_program():
    root.destroy()

root = tk.Tk()
root.title("Upload Excel File")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

label_select = tk.Label(frame, text="Select Excel File:")
label_select.pack(side=tk.TOP, anchor='w')

entry = tk.Entry(frame, width=40)
entry.pack(side=tk.LEFT, padx=5)

browse_button = tk.Button(frame, text="Browse", command=select_file)
browse_button.pack(side=tk.LEFT)

upload_button = tk.Button(root, text="Upload", command=upload_file)
upload_button.pack(side=tk.LEFT, padx=5, pady=5)

exit_button = tk.Button(root, text="Exit", command=exit_program)
exit_button.pack(side=tk.LEFT, padx=5, pady=5)

root.mainloop()
