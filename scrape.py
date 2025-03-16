import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# Function to parse XML and extract data into a DataFrame
def parse_pdbxml(file_name, encoding='utf-8', namespace=None):
    try:
        tree = ET.parse(file_name, parser=ET.XMLParser(encoding=encoding))
        root = tree.getroot()
    except Exception as e:
        messagebox.showerror("Error", f"Error loading XML file: {e}")
        return None
    
    data = []
    ns = { 'pdb': namespace } if namespace else {}

    for form in root.findall("pdb:form", ns):
        form_name = form.get("name")
        for test in form.findall("pdb:test", ns):
            row = {
                "form_name": form_name,
                "test_date": test.get("date"),
                "resultsguid": test.get("resultsguid"),
            }
            for tag in test.find("pdb:data", ns).findall("pdb:tag", ns):
                row[tag.get("name")] = tag.text if tag.text is not None else ""
            for array in test.find("pdb:data", ns).findall("pdb:array", ns):
                array_name = array.get("name")
                row[array_name] = ", ".join(
                    item.text for item in array.findall("pdb:arrayitem", ns) if item.text is not None
                )
            data.append(row)

    return data

# Function to open the file dialog and set the file path
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDBXML Files", "*.pdbxml")])
    file_entry.delete(0, tk.END)
    file_entry.insert(0, file_path)

# Function to generate the CSV from the selected XML
def generate_csv():
    file_path = file_entry.get()
    encoding = encoding_entry.get() if encoding_entry.get() else "utf-8"
    namespace = namespace_entry.get()

    if not file_path:
        messagebox.showwarning("Input Error", "Please select a PDBXML file.")
        return
    
    data = parse_pdbxml(file_path, encoding, namespace)
    if data:
        df = pd.DataFrame(data)
        output_file = os.path.splitext(file_path)[0] + "_output.csv"
        df.to_csv(output_file, index=False)
        messagebox.showinfo("Success", f"CSV file generated: {output_file}")

# Set up the Tkinter window
window = tk.Tk()
window.title("PDBXML to CSV Converter")

# UI elements
tk.Label(window, text="Select PDBXML File:").pack(pady=10)
file_entry = tk.Entry(window, width=40)
file_entry.pack(pady=5)
tk.Button(window, text="Browse", command=browse_file).pack(pady=5)

tk.Label(window, text="Encoding (Default: utf-8):").pack(pady=10)
encoding_entry = tk.Entry(window, width=40)
encoding_entry.pack(pady=5)

tk.Label(window, text="Namespace (Optional):").pack(pady=10)
namespace_entry = tk.Entry(window, width=40)
namespace_entry.pack(pady=5)

tk.Button(window, text="Generate CSV", command=generate_csv).pack(pady=20)

# Run the Tkinter event loop
window.mainloop()