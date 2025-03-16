import argparse
import xml.etree.ElementTree as ET
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

def parse_xml(file_path):
    """Parses an XML or PDBXML file and extracts structured data for each battery cell."""
    tree = ET.parse(file_path)
    root = tree.getroot()

    # Extract general test details
    general_info = {}
    for form in root.findall("form"):
        general_info["Form Name"] = form.get("name")
        for test in form.findall("test"):
            general_info["Test Date"] = test.get("date")
            general_info["Results GUID"] = test.get("resultsguid")

            # Extract per-cell data (organized horizontally)
            cell_data = []
            for array in test.findall(".//array"):
                array_name = array.get("name")

                # Process array values grouped by cell number
                for item in array.findall("arrayitem"):
                    cell_no = int(item.get("index"))
                    value = item.text if item.text is not None else ""

                    # Find or create entry for this cell
                    cell_entry = next((cell for cell in cell_data if cell["Cell No"] == cell_no), None)
                    if not cell_entry:
                        cell_entry = {"Cell No": cell_no}
                        cell_data.append(cell_entry)

                    cell_entry[array_name] = value

            return general_info, cell_data

def write_excel(general_info, cell_data, output_file):
    """Writes extracted data into a well-structured Excel (.xlsx) file."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Battery Test Report"

    # Formatting Headers
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center")

    # 1️⃣ **General Test Information Section**
    ws.append(["Battery Test Report"])
    ws["A1"].font = bold_font
    ws.append([])
    
    for key, value in general_info.items():
        ws.append([key, value])

    ws.append([])  # Add empty row for spacing

    # 2️⃣ **Cell Data Table**
    headers = ["Cell No", "Impedance (mΩ)", "Voltage (V)", "Temperature (°C)", "Specific Gravity", "Time"]
    ws.append(headers)

    for col in ["A", "B", "C", "D", "E", "F"]:
        ws["{}{}".format(col, ws.max_row)].font = bold_font
        ws["{}{}".format(col, ws.max_row)].alignment = center_align

    # Write cell data
    for row in cell_data:
        ws.append([
            row.get("Cell No", ""),
            row.get("impedance", ""),
            row.get("voltage", ""),
            row.get("temp", ""),
            row.get("specific_gravity", ""),
            row.get("time", "")
        ])

    # Adjust column widths
    for col in ws.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 2

    # Save workbook
    wb.save(output_file)

def main():
    parser = argparse.ArgumentParser(description="Convert PDBXML/XML to Excel (.xlsx) with structured formatting")
    parser.add_argument("input_file", help="Path to the XML/PDBXML file")
    parser.add_argument("-o", "--output", help="Output Excel file name (default: input_file.xlsx)", default=None)

    args = parser.parse_args()
    input_file = args.input_file
    output_file = args.output or os.path.splitext(input_file)[0] + "_report.xlsx"

    if not os.path.exists(input_file):
        print(f"Error: File '{input_file}' not found.")
        return

    print(f"Processing: {input_file}")
    general_info, cell_data = parse_xml(input_file)
    write_excel(general_info, cell_data, output_file)
    
    print(f"✅ Conversion complete! Excel file saved as: {output_file}")

if __name__ == "__main__":
    main()