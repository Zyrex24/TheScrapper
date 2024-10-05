import requests
import openpyxl
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog


# Function to fetch data based on user input and save it to an Excel file
def fetch_data():
    # Get values from the GUI
    url_template = url_template_entry.get()
    bearer_token = bearer_token_entry.get()
    student_ids = ids_entry.get().split(',')

    headers = {
        'Authorization': f'Bearer {bearer_token}'
    }

    # Create a new workbook and select the active worksheet
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Student Courses"

    # Create headers for the columns dynamically based on selections
    selected_fields = []
    if var_id.get(): selected_fields.append("Student ID")
    if var_name.get(): selected_fields.append("Student Name")
    if var_gpa.get(): selected_fields.append("GPA")
    if var_program.get(): selected_fields.append("Program")
    if var_level.get(): selected_fields.append("Level")
    if var_courses.get(): selected_fields.append("Total Courses")

    # Create headers for the columns
    sheet.append(selected_fields)

    # Iterate over the list of IDs and make a request for each one
    for student_id in student_ids:
        try:
            url = url_template.format(id=student_id)  # Replace {id} with the actual ID
            response = requests.get(url, headers=headers)

            # Convert the response to JSON (a list of courses for the student)
            data = response.json()

            # Check if data is available
            if data:
                row_data = []
                if var_id.get(): row_data.append(data[0]['student']['id'])
                if var_name.get(): row_data.append(data[0]['student']['name'])
                if var_gpa.get(): row_data.append(data[0]['student']['gpa'])
                if var_program.get(): row_data.append(data[0]['student']['program']['name'])
                if var_level.get(): row_data.append(data[0]['student']['level']['name'])
                if var_courses.get(): row_data.append(len(data))

                # Append the row to the Excel sheet
                sheet.append(row_data)

        except Exception as e:
            print(f"Failed to retrieve data for student ID {student_id}: {e}")
            continue

    # Save the Excel file
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        wb.save(save_path)
        messagebox.showinfo("Success", "Data has been written to the Excel file.")
    else:
        messagebox.showerror("Error", "No file path selected for saving.")


# Create the GUI window
root = tk.Tk()
root.title("Student Data Fetcher")

# URL Template Label and Entry
tk.Label(root, text="URL Template:").grid(row=0, column=0, padx=10, pady=10)
url_template_entry = tk.Entry(root, width=50)
url_template_entry.grid(row=0, column=1, padx=10, pady=10)

# Bearer Token Label and Entry
tk.Label(root, text="Bearer Token:").grid(row=1, column=0, padx=10, pady=10)
bearer_token_entry = tk.Entry(root, width=50)
bearer_token_entry.grid(row=1, column=1, padx=10, pady=10)

# IDs Label and Entry
tk.Label(root, text="Student IDs (comma-separated):").grid(row=2, column=0, padx=10, pady=10)
ids_entry = tk.Entry(root, width=50)
ids_entry.grid(row=2, column=1, padx=10, pady=10)

# Checkboxes to customize what to fetch
var_id = tk.BooleanVar()
var_name = tk.BooleanVar()
var_gpa = tk.BooleanVar()
var_program = tk.BooleanVar()
var_level = tk.BooleanVar()
var_courses = tk.BooleanVar()

tk.Checkbutton(root, text="Student ID", variable=var_id).grid(row=3, column=0, sticky="w")
tk.Checkbutton(root, text="Student Name", variable=var_name).grid(row=3, column=1, sticky="w")
tk.Checkbutton(root, text="GPA", variable=var_gpa).grid(row=4, column=0, sticky="w")
tk.Checkbutton(root, text="Program", variable=var_program).grid(row=4, column=1, sticky="w")
tk.Checkbutton(root, text="Level", variable=var_level).grid(row=5, column=0, sticky="w")
tk.Checkbutton(root, text="Total Courses", variable=var_courses).grid(row=5, column=1, sticky="w")

# Fetch Data Button
fetch_button = tk.Button(root, text="Fetch Data", command=fetch_data)
fetch_button.grid(row=6, column=0, columnspan=2, pady=20)

# Start the GUI event loop
root.mainloop()
