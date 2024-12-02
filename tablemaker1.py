import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import json
import PyPDF2
import subprocess

class MappingTool:
    def __init__(self, root, json_data):
        self.root = root
        self.json_data = json_data
        self.mappings = {}
        self.all_fields = []  # To store all fields for searching

        self.setup_gui()

    def setup_gui(self):
        # Create the Notebook (tabs container)
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True)

        # Create frames for each tab
        self.mapping_tab = ttk.Frame(notebook)
        self.pdf_writer_tab = ttk.Frame(notebook)
        self.ocr_to_json_tab = ttk.Frame(notebook)

        # Add tabs to the notebook
        notebook.add(self.mapping_tab, text="JSON PDF Mapping")
        notebook.add(self.pdf_writer_tab, text="PDF Writer")
        notebook.add(self.ocr_to_json_tab, text="OCR to JSON")

        # Add the existing mapping code to the JSON PDF Mapping tab
        self.create_mapping_tab()

        # Add the PDF Writer tab functionality
        self.create_pdf_writer_tab()

    def create_mapping_tab(self):
        """Creates the content for the 'JSON PDF Mapping' tab."""
        # Main frame for JSON PDF Mapping
        main_frame = ttk.Frame(self.mapping_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Top frame to hold both sides
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill=tk.BOTH, expand=True)

        # Left frame for JSON structure
        left_frame = ttk.Frame(top_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Label for JSON structure tree
        self.json_structure_label = ttk.Label(left_frame, text="Input JSON Lookup Structure:")
        self.json_structure_label.pack(fill=tk.X)

        # JSON Tree View
        self.tree = ttk.Treeview(left_frame)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.populate_tree(self.tree, '', self.json_data)

        # Right frame to hold PDF file input, form field selection, type selector, and mapping list
        right_frame = ttk.Frame(top_frame, padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Input PDF Form Filename
        self.pdf_file_label = ttk.Label(right_frame, text="Input PDF Form Filename:")
        self.pdf_file_label.pack(fill=tk.X)
        self.pdf_file_var = tk.StringVar(self.root)
        self.pdf_file_entry = ttk.Entry(right_frame, textvariable=self.pdf_file_var)
        self.pdf_file_entry.pack(fill=tk.X, pady=5)
        
        # Load PDF Button
        self.load_button = ttk.Button(right_frame, text="Load PDF", command=self.load_pdf_fields)
        self.load_button.pack(fill=tk.X, pady=5)

        # Search Box
        self.search_label = ttk.Label(right_frame, text="Search Form Fields:")
        self.search_label.pack(fill=tk.X)
        self.search_var = tk.StringVar(self.root)
        self.search_var.trace_add("write", self.update_search_results)
        self.search_entry = ttk.Entry(right_frame, textvariable=self.search_var)
        self.search_entry.pack(fill=tk.X, pady=5)

        # Available Form Fields Tree View
        self.field_label = ttk.Label(right_frame, text="Available Form Fields:")
        self.field_label.pack(fill=tk.X)
        self.fields_tree = ttk.Treeview(right_frame, columns=('Field',), show='tree')
        self.fields_tree.pack(fill=tk.BOTH, expand=True, pady=5)

        # Form Type Selector
        self.type_label = ttk.Label(right_frame, text="Form Field Type:")
        self.type_label.pack(fill=tk.X)
        self.type_var = tk.StringVar(self.root)
        self.type_selector = ttk.Combobox(right_frame, textvariable=self.type_var)
        self.type_selector['values'] = ("FILL_FIELD", "FILL_ADDRESS", "CHECKBOX", "RADIO_BUTTON")
        self.type_selector.pack(fill=tk.X, pady=5)

        # Allowed Values Entry for Checkboxes and Radio Buttons
        self.allowed_values_label = ttk.Label(right_frame, text='Allowed Values (for CHECKBOX/RADIO_BUTTON):\nFormat: ["Value 1", "Value 2", "Value n"]')
        self.allowed_values_label.pack(fill=tk.X)
        self.allowed_values_var = tk.StringVar(self.root)
        self.allowed_values_entry = ttk.Entry(right_frame, textvariable=self.allowed_values_var)
        self.allowed_values_entry.pack(fill=tk.X, pady=5)

        # Map Button
        self.map_button = ttk.Button(right_frame, text="Map", command=self.add_mapping)
        self.map_button.pack(fill=tk.X, pady=5)

        # Output File Name
        self.output_file_label = ttk.Label(right_frame, text="Output Lookup Table File:")
        self.output_file_label.pack(fill=tk.X)
        self.output_file_var = tk.StringVar(self.root)
        self.output_file_entry = ttk.Entry(right_frame, textvariable=self.output_file_var)
        self.output_file_entry.pack(fill=tk.X, pady=5)

        # Save Button
        self.save_button = ttk.Button(right_frame, text="Save Mappings", command=self.save_mappings)
        self.save_button.pack(fill=tk.X, pady=5)

        # Mapping List
        self.mappings_list_label = ttk.Label(right_frame, text="Mappings:")
        self.mappings_list_label.pack(fill=tk.X)
        self.mappings_list = tk.Listbox(right_frame)
        self.mappings_list.pack(fill=tk.BOTH, expand=True)

    def create_pdf_writer_tab(self):
        """Creates the content for the 'PDF Writer' tab."""
        main_frame = ttk.Frame(self.pdf_writer_tab, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # PDF Input File Dropdown
        self.pdf_input_label = ttk.Label(main_frame, text="Input - Fillable pdf form filename:")
        self.pdf_input_label.pack(fill=tk.X)
        self.pdf_input_var = tk.StringVar(self.root)
        self.pdf_input_combobox = ttk.Combobox(main_frame, textvariable=self.pdf_input_var)
        self.populate_combobox(self.pdf_input_combobox, '.pdf')
        self.pdf_input_combobox.pack(fill=tk.X, pady=5)

        # PDF Output File Dropdown
        self.pdf_output_label = ttk.Label(main_frame, text="Output - filled pdf form filename:")
        self.pdf_output_label.pack(fill=tk.X)
        self.pdf_output_var = tk.StringVar(self.root)
        self.pdf_output_combobox = ttk.Combobox(main_frame, textvariable=self.pdf_output_var)
        self.populate_combobox(self.pdf_output_combobox, '.pdf')
        self.pdf_output_combobox.pack(fill=tk.X, pady=5)

        # JSON Data File Dropdown
        self.json_data_label = ttk.Label(main_frame, text="Input - provider data json filename:")
        self.json_data_label.pack(fill=tk.X)
        self.json_data_var = tk.StringVar(self.root)
        self.json_data_combobox = ttk.Combobox(main_frame, textvariable=self.json_data_var)
        self.populate_combobox(self.json_data_combobox, '.json')
        self.json_data_combobox.pack(fill=tk.X, pady=5)

        # Lookup Table JSON File Dropdown
        self.lookup_table_label = ttk.Label(main_frame, text="Input - json-pdf lookup filename:")
        self.lookup_table_label.pack(fill=tk.X)
        self.lookup_table_var = tk.StringVar(self.root)
        self.lookup_table_combobox = ttk.Combobox(main_frame, textvariable=self.lookup_table_var)
        self.populate_combobox(self.lookup_table_combobox, '.json')
        self.lookup_table_combobox.pack(fill=tk.X, pady=5)

        # Run PDF Writer Script Button
        self.run_button = ttk.Button(main_frame, text="Run PDF Writer", command=self.run_pdf_writer_script)
        self.run_button.pack(fill=tk.X, pady=5)

    def populate_combobox(self, combobox, file_extension):
        """Populates the dropdown with files from the current directory."""
        current_directory = os.getcwd()
        files = [f for f in os.listdir(current_directory) if f.endswith(file_extension)]
        combobox['values'] = files

    def run_pdf_writer_script(self):
        """Calls the external Python script with selected parameters."""
        input_pdf = self.pdf_input_var.get()
        output_pdf = self.pdf_output_var.get()
        json_data = self.json_data_var.get()
        lookup_table = self.lookup_table_var.get()

        if not (input_pdf and output_pdf and json_data and lookup_table):
            messagebox.showerror("Error", "Please provide all parameters.")
            return

        # Call the external script with parameters
        try:
            subprocess.run(
                ["python3", "priv_look_work1.py", input_pdf, output_pdf, json_data, lookup_table],
                check=True
            )
            messagebox.showinfo("Success", "PDF Writer script executed successfully!")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Script failed: {e}")

    def populate_tree(self, tree, parent, data):
        """Recursively populates the tree with JSON data"""
        if isinstance(data, dict):
            for key, value in data.items():
                node = tree.insert(parent, 'end', text=key, open=True)
                self.populate_tree(tree, node, value)
        elif isinstance(data, list):
            for index, item in enumerate(data):
                node = tree.insert(parent, 'end', text=f"[{index}]", open=True)  # Fixed this line
                self.populate_tree(tree, node, item)

    def load_pdf_fields(self):
        """Load the fields from the specified PDF file"""
        pdf_file_name = self.pdf_file_var.get()
        try:
            with open(pdf_file_name, 'rb') as pdf_file:
                reader = PyPDF2.PdfReader(pdf_file)
                if "/AcroForm" in reader.trailer['/Root']:
                    form_fields = reader.trailer['/Root']['/AcroForm']['/Fields']
                    fields = [field.get_object().get('/T') for field in form_fields]
                    self.all_fields = [field for field in fields if field is not None]
                    self.populate_fields_tree(self.fields_tree, self.all_fields)
                else:
                    messagebox.showerror("Error", "No form fields found in the PDF.")
        except FileNotFoundError:
            messagebox.showerror("Error", "PDF file not found.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def populate_fields_tree(self, tree, fields):
        """Populates the fields tree with available PDF form fields"""
        tree.delete(*tree.get_children())  # Clear the tree first
        for field in fields:
            tree.insert('', 'end', text=field)

    def update_search_results(self, *args):
        """Update the fields tree based on the search query"""
        search_query = self.search_var.get().lower()
        filtered_fields = [field for field in self.all_fields if search_query in field.lower()]
        self.populate_fields_tree(self.fields_tree, filtered_fields)

    def add_mapping(self):
        """Adds the selected JSON path and form field to the mapping list"""
        selected_json_item = self.tree.selection()[0]
        json_path = self.get_full_path(self.tree, selected_json_item)

        selected_field_item = self.fields_tree.selection()[0]
        pdf_field = self.fields_tree.item(selected_field_item, "text")

        field_type = self.type_var.get()
        allowed_values = self.allowed_values_var.get()

        if pdf_field and field_type:
            mapping = {
                "json_path": json_path,
                "type": field_type
            }
            if field_type in ("CHECKBOX", "RADIO_BUTTON") and allowed_values:
                try:
                    allowed_values_list = json.loads(allowed_values)  # Try to parse the entered allowed values
                    mapping["allowed_values"] = allowed_values_list
                    self.mappings_list.insert(tk.END, f"{pdf_field} -> {json_path} ({field_type}) [Allowed: {allowed_values_list}]")
                except json.JSONDecodeError:
                    messagebox.showerror("Error", "Invalid allowed values format. Please enter values in the correct format.")
                    return
            else:
                self.mappings_list.insert(tk.END, f"{pdf_field} -> {json_path} ({field_type})")

            # Add mapping to the dictionary
            self.mappings[pdf_field] = mapping

            # Clear the form field type and allowed values entry after adding the mapping
            self.type_var.set("")
            self.allowed_values_var.set("")

    def get_full_path(self, tree, item):
        """Gets the full path of the selected JSON element"""
        path = []
        while item:
            path.append(tree.item(item, "text"))
            item = tree.parent(item)
        path.reverse()
        return " -> ".join(path)

    def save_mappings(self):
        """Save the current mappings to the specified output file in JSON format"""
        output_file_name = self.output_file_var.get()
        try:
            with open(output_file_name, 'r+') as f:
                existing_data = json.load(f)
                existing_data.update(self.mappings)
                f.seek(0)
                json.dump(existing_data, f, indent=2)
                f.truncate()
            messagebox.showinfo("Success", "Mappings saved successfully!")
            # Clear mappings after saving
            self.mappings.clear()
            self.mappings_list.delete(0, tk.END)  # Clear the mappings listbox
        except FileNotFoundError:
            with open(output_file_name, 'w') as f:
                json.dump(self.mappings, f, indent=2)
            messagebox.showinfo("Success", "Mappings saved successfully!")
            # Clear mappings after saving
            self.mappings.clear()
            self.mappings_list.delete(0, tk.END)  # Clear the mappings listbox
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Load JSON data
with open('data.json') as f:
    json_data = json.load(f)

root = tk.Tk()
root.title("PDF Form Mapping Tool")
root.geometry("800x600")
app = MappingTool(root, json_data)
root.mainloop()
