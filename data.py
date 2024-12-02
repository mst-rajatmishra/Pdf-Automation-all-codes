import json
import os
import fitz  # PyMuPDF
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Combobox

# Function to load JSON data from a file
def load_json(file_path):
    try:
        with open(file_path, 'r') as file:
            return json.load(file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load JSON file: {e}")
        return None

# Function to load the lookup table (JSON file)
def load_lookup_table(file_path):
    try:
        with open(file_path, 'r') as file:
            return json.load(file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load lookup table: {e}")
        return None

# Extract values from JSON data based on the given path (e.g., "PERSONAL INFORMATION -> Name -> First Name :")
def extract_value_from_json(data, path, default_value=None):
    keys = path.split(' -> ')
    for key in keys:
        try:
            if isinstance(data, list):
                index = int(key.strip("[]"))
                data = data[index]
            else:
                data = data.get(key, default_value)
        except (KeyError, IndexError, ValueError) as e:
            return default_value
    return data

# Function to fill the PDF form fields based on JSON data and lookup table
def fill_pdf(input_pdf_path, output_pdf_path, json_data, lookup_table, password=None):
    try:
        pdf_document = fitz.open(input_pdf_path)
        for page in pdf_document:
            for widget in page.widgets():
                field_name = widget.field_name
                if field_name in lookup_table:
                    json_path = lookup_table[field_name]['json_path']
                    field_type = lookup_table[field_name]['type']
                    field_value = extract_value_from_json(json_data, json_path, "N/A")
                    
                    if field_type == "FILL_FIELD":
                        if field_value:
                            widget.field_value = field_value
                    elif field_type == "FILL_ADDRESS":
                        if isinstance(field_value, dict):
                            address = ', '.join(filter(None, [
                                field_value.get("Street1 :"),
                                field_value.get("Street2 :"),
                                field_value.get("City :"),
                                field_value.get("State :"),
                                field_value.get("Zip Code :")
                            ]))
                            widget.field_value = address
                    elif field_type == "CHECKBOX":
                        allowed_values = lookup_table[field_name].get('allowed_values', [])
                        if field_value in allowed_values:
                            widget.field_value = True if field_value else False
                    elif field_type == "RADIO_BUTTON":
                        allowed_values = lookup_table[field_name].get('allowed_values', [])
                        if field_value in allowed_values:
                            widget.field_value = True

                    widget.update()

        # Save the filled PDF, applying password protection if provided
        if password:
            pdf_document.save(output_pdf_path, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=password)
        else:
            pdf_document.save(output_pdf_path)

        pdf_document.close()

    except Exception as e:
        messagebox.showerror("Error", f"Error while processing PDF: {e}")

# Function to check if a file already exists
def check_pdf_exists(file_path):
    return os.path.exists(file_path)

# Function to process all users (or single user) and generate PDFs
def generate_pdfs(input_pdf_path, json_data, lookup_table, output_dir, progress):
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        existing_users = set()
        if isinstance(json_data, list):
            progress['maximum'] = len(json_data)
            for idx, user_data in enumerate(json_data):
                first_name = extract_value_from_json(user_data, "PERSONAL INFORMATION -> Name -> First Name :", "Unknown")
                last_name = extract_value_from_json(user_data, "PERSONAL INFORMATION -> Name -> Last Name :", "User")
                output_pdf_file = os.path.join(output_dir, f"output_{first_name}_{last_name}_{idx + 1}.pdf")

                # Skip if file already exists or data has been processed before
                if check_pdf_exists(output_pdf_file):
                    messagebox.showwarning("Warning", f"PDF for {first_name} {last_name} already exists. Skipping.")
                    continue
                user_identifier = f"{first_name}_{last_name}"
                if user_identifier in existing_users:
                    messagebox.showwarning("Warning", f"Data for {first_name} {last_name} already processed.")
                    continue

                fill_pdf(input_pdf_path, output_pdf_file, user_data, lookup_table)
                existing_users.add(user_identifier)

                # Update progress
                progress['value'] += 1
                root.update_idletasks()

            messagebox.showinfo("Success", "All PDFs processed successfully!")

        else:
            output_pdf_file = os.path.join(output_dir, "output_single_user.pdf")
            if check_pdf_exists(output_pdf_file):
                messagebox.showwarning("Warning", "PDF already exists. Skipping.")
                return

            fill_pdf(input_pdf_path, output_pdf_file, json_data, lookup_table)
            messagebox.showinfo("Success", "PDF generated successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to open file dialog and select a file
def select_file(entry_widget, file_types, path_var):
    file_path = filedialog.askopenfilename(filetypes=file_types)
    if file_path:
        entry_widget.delete(0, END)
        entry_widget.insert(0, os.path.basename(file_path))  # Display file name only
        path_var.set(file_path)  # Store full path in a variable

# Function to open directory dialog and select output folder
def select_directory(entry_widget):
    directory_path = filedialog.askdirectory()
    if directory_path:
        entry_widget.delete(0, END)
        entry_widget.insert(0, os.path.basename(directory_path))
    return directory_path

# Function to get file types based on selected database type
def get_file_types(database_type):
    if database_type == "JSON":
        return [("JSON files", "*.json")]
    elif database_type == "CSV":
        return [("CSV files", "*.csv")]
    elif database_type == "XML":
        return [("XML files", "*.xml")]
    return []

# Create the main GUI window
root = Tk()
root.title("PDF Form Filler")
root.geometry("500x500")

# Variables to hold full file paths
pdf_path_var = StringVar()
json_path_var = StringVar()
lookup_path_var = StringVar()
database_type_var = StringVar(value="JSON")  # Default to JSON

# PDF Template Selection
Label(root, text="Select PDF Template").pack(pady=5)
pdf_entry = Entry(root, width=50, textvariable=pdf_path_var)
pdf_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_file(pdf_entry, [("PDF files", "*.pdf")], pdf_path_var)).pack(pady=5)

# Database Type Dropdown
Label(root, text="Select Database Type").pack(pady=5)
database_type_dropdown = Combobox(root, textvariable=database_type_var, values=["JSON", "XML", "CSV"])
database_type_dropdown.pack(pady=5)

# Database File Selection
Label(root, text="Select Database File").pack(pady=5)
json_entry = Entry(root, width=50, textvariable=json_path_var)
json_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_file(json_entry, get_file_types(database_type_var.get()), json_path_var)).pack(pady=5)

# Lookup Table Selection
Label(root, text="Select Lookup Table").pack(pady=5)
lookup_entry = Entry(root, width=50, textvariable=lookup_path_var)
lookup_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_file(lookup_entry, [("JSON files", "*.json")], lookup_path_var)).pack(pady=5)

# Output Directory Selection
Label(root, text="Select Output Directory").pack(pady=5)
output_entry = Entry(root, width=50)
output_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_directory(output_entry)).pack(pady=5)

# Progress Bar
progress = Progressbar(root, orient=HORIZONTAL, length=400, mode='determinate')
progress.pack(pady=20)

# Process Button
def process_pdfs():
    pdf_path = pdf_path_var.get()
    json_path = json_path_var.get()
    lookup_path = lookup_path_var.get()
    output_dir = output_entry.get()

    # Load the JSON data and lookup table
    json_data = load_json(json_path)
    if not json_data:
        return
    lookup_table = load_lookup_table(lookup_path)
    if not lookup_table:
        return

    # Generate PDFs
    generate_pdfs(pdf_path, json_data, lookup_table, output_dir, progress)

Button(root, text="Process", command=process_pdfs).pack(pady=20)

# Run the GUI
root.mainloop()
