import json
import fitz  # PyMuPDF
import os
from tkinter import *
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Combobox
from tqdm import tqdm

# Load the JSON data
def load_json(json_file):
    with open(json_file, 'r') as file:
        return json.load(file)

# Load the lookup table
def load_lookup_table(lookup_file):
    with open(lookup_file, 'r') as file:
        return json.load(file)

# Extract value from JSON using the JSON path
def extract_value_from_json(json_data, json_path, default_value=None):
    keys = json_path.split(' -> ')
    for key in keys:
        if isinstance(json_data, list):
            try:
                key = int(key[1:-1]) if key.startswith('[') and key.endswith(']') else int(key)
                json_data = json_data[key]
            except (ValueError, IndexError) as e:
                print(f"Error accessing key '{key}': {e}")
                return default_value
        else:
            json_data = json_data.get(key, default_value)
    return json_data

# Fill PDF form fields using data from JSON and lookup table
def fill_pdf(input_pdf_file, output_pdf_file, json_data, lookup_table, password=None):
    try:
        pdf_document = fitz.open(input_pdf_file)
        for page_num in range(pdf_document.page_count):
            page = pdf_document.load_page(page_num)
            for field in page.widgets():
                field_name = field.field_name
                if field_name in lookup_table:
                    json_path = lookup_table[field_name]['json_path']
                    field_type = lookup_table[field_name]['type']
                    field_value = extract_value_from_json(json_data, json_path, default_value="N/A")
                    
                    if field_type == "FILL_FIELD":
                        if field_value:
                            field.field_value = field_value
                            field.update()
                    elif field_type == "FILL_ADDRESS":
                        if isinstance(field_value, dict):
                            address_parts = [
                                field_value.get("Street1 :"),
                                field_value.get("Street2 :"),
                                field_value.get("City :"),
                                field_value.get("State :"),
                                field_value.get("Zip Code :")
                            ]
                            address = ', '.join([part for part in address_parts if part])
                            field.field_value = address
                            field.update()
                    elif field_type == "CHECKBOX":
                        allowed_values = lookup_table[field_name]['allowed_values']
                        if field_value and field_value in allowed_values:
                            field.field_value = field_value
                            field.update()
                    elif field_type == "RADIO_BUTTON":
                        allowed_values = lookup_table[field_name]['allowed_values']
                        if field_value and field_value.upper() in allowed_values:
                            # For sex radio buttons
                            if field_name == "sex":  # Ensure this matches your field name in the PDF
                                if field_value.upper() == "MALE":
                                    field.field_value = "Male"  # Match the exact field value in the PDF
                                elif field_value.upper() == "FEMALE":
                                    field.field_value = "Female"  # Match the exact field value in the PDF
                            else:
                                field.field_value = "Yes" if field_value.upper() == "YES" else "No"
                            field.update()

        # Save the PDF with optional password protection
        if password:
            pdf_document.save(output_pdf_file, encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=password)
        else:
            pdf_document.save(output_pdf_file)
        
        pdf_document.close()
    except Exception as e:
        print(f"Error while filling PDF: {e}")

# Function to check if PDF already exists
def check_pdf_exists(output_pdf_file):
    return os.path.exists(output_pdf_file)

# Function to process all users and generate PDFs
def generate_pdfs(input_pdf_file, json_data_file, lookup_table_file, output_dir, progress):
    try:
        # Load data from JSON and lookup table
        json_data = load_json(json_data_file)
        lookup_table = load_lookup_table(lookup_table_file)

        # Ensure the output directory exists
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        existing_users = set()
        if isinstance(json_data, list):
            progress['maximum'] = len(json_data)
            for idx, user_data in enumerate(json_data):
                first_name = extract_value_from_json(user_data, "PERSONAL INFORMATION -> Name -> First Name :", "Unknown")
                last_name = extract_value_from_json(user_data, "PERSONAL INFORMATION -> Name -> Last Name :", "User")
                user_identifier = f"{first_name}_{last_name}"
                output_pdf_file = os.path.join(output_dir, f"output_{first_name}_{last_name}_{idx + 1}.pdf")

                if check_pdf_exists(output_pdf_file):
                    messagebox.showwarning("Warning", f"PDF for {first_name} {last_name} already exists.")
                    continue
                
                if user_identifier in existing_users:
                    messagebox.showwarning("Warning", f"Data for {first_name} {last_name} already processed.")
                    continue

                fill_pdf(input_pdf_file, output_pdf_file, user_data, lookup_table)
                existing_users.add(user_identifier)
                progress['value'] += 1
                root.update_idletasks()
            messagebox.showinfo("Success", "PDFs processed successfully!")
        else:
            output_pdf_file = os.path.join(output_dir, "output_single_user.pdf")
            if check_pdf_exists(output_pdf_file):
                messagebox.showwarning("Warning", "PDF already exists.")
                return
            
            fill_pdf(input_pdf_file, output_pdf_file, json_data, lookup_table)
            messagebox.showinfo("Success", "PDF generated successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Function to open file dialog and select a file
def select_file(entry_widget, file_types, full_path_var):
    file_path = filedialog.askopenfilename(filetypes=file_types)
    if file_path:
        entry_widget.delete(0, END)
        entry_widget.insert(0, os.path.basename(file_path))  # Show only the file name
        full_path_var.set(file_path)  # Store the full path in a separate variable

# Function to open directory dialog and select output folder
def select_directory(entry_widget):
    directory_path = filedialog.askdirectory()
    entry_widget.delete(0, END)
    entry_widget.insert(0, os.path.basename(directory_path))  # Corrected this line to use entry_widget
    return directory_path

# Create the GUI
root = Tk()
root.title("PDF Form Filler")
root.geometry("500x400")

# Variables to hold full file paths
pdf_path_var = StringVar()
json_path_var = StringVar()
lookup_path_var = StringVar()
database_type_var = StringVar(value="JSON")  # Default to JSON

# Input PDF file
Label(root, text="Select PDF Template").pack(pady=5)
pdf_entry = Entry(root, width=50)
pdf_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_file(pdf_entry, [("PDF files", "*.pdf")], pdf_path_var)).pack(pady=5)

# Dropdown for database type
Label(root, text="Select Database Type").pack(pady=5)
database_type_dropdown = Combobox(root, textvariable=database_type_var, values=["JSON", "XML", "CSV"])
database_type_dropdown.pack(pady=5)

# Database file
Label(root, text="Select Database File").pack(pady=5)
json_entry = Entry(root, width=50)
json_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_file(json_entry, get_file_types(), json_path_var)).pack(pady=5)

# Lookup Table file
Label(root, text="Select Lookup Table").pack(pady=5)
lookup_entry = Entry(root, width=50)
lookup_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_file(lookup_entry, [("JSON files", "*.json")], lookup_path_var)).pack(pady=5)

# Output Directory
Label(root, text="Select Output Directory").pack(pady=5)
output_entry = Entry(root, width=50)
output_entry.pack(pady=5)
Button(root, text="Browse", command=lambda: select_directory(output_entry)).pack(pady=5)

# Progress bar
progress = Progressbar(root, orient=HORIZONTAL, length=400, mode='determinate')
progress.pack(pady=20)

# Function to get file types based on selected database type
def get_file_types():
    db_type = database_type_var.get()
    if db_type == "JSON":
        return [("JSON files", "*.json")]
    elif db_type == "CSV":
        return [("CSV files", "*.csv")]
    elif db_type == "XML":
        return [("XML files", "*.xml")]
    return []

# Process button
Button(root, text="Process", command=lambda: generate_pdfs(pdf_path_var.get(), json_path_var.get(), lookup_path_var.get(), output_entry.get(), progress)).pack(pady=20)

# Run the GUI
root.mainloop()
