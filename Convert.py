import pdfplumber
import pandas as pd
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import os

# Input PDF
input_pdf = "Consolidated Train Performance_05_08_2022.pdf"

all_tables = []

# Define base SharePoint sync path
sharepoint_root = r"C:\Users\Guillaume Turillon\RATP Dev\PowerAutomateAnalysis - Analysis"  # Change if needed

# Extract date and tables from PDF
with pdfplumber.open(input_pdf) as pdf:
    first_page_text = pdf.pages[0].extract_text()
    match = re.search(r"Date\s*:\s*(\d{4})-(\d{2})-(\d{2})", first_page_text)
    
    if match:
        year, month, day = match.groups()

        # Build dynamic output folder path and ensure it exists
        output_folder = os.path.join(sharepoint_root, year, month, day)
        os.makedirs(output_folder, exist_ok=True)

        # Define full output Excel path
        output_excel = os.path.join(output_folder, f"QOS_{day}_{month}_{year}.xlsx")
    else:
        raise ValueError("❌ Could not extract date from PDF. Make sure the date format is 'Date : YYYY-MM-DD'.")

    # Extract tables from all pages
    for i, page in enumerate(pdf.pages):
        print(f"Processing page {i+1}/{len(pdf.pages)}")
        tables = page.extract_tables()
        for table in tables:
            df = pd.DataFrame(table[1:], columns=table[0])
            df = df.applymap(lambda x: x.strip().replace('\u200b', '') if isinstance(x, str) else x)
            for col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='ignore')
            all_tables.append(df)

# Combine all data into one DataFrame
final_df = pd.concat(all_tables, ignore_index=True)

# Create workbook and worksheet
wb = Workbook()
ws = wb.active
ws.title = "QOS Data"

# Define styles
header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")  # Header: Dark blue
light_row_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")  # Light blue
dark_row_fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")  # Darker blue
header_font = Font(bold=True, color="FFFFFF")  # White font
center_alignment = Alignment(horizontal="center", vertical="center")

# Write header row with style
for col_idx, col_name in enumerate(final_df.columns, start=1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = center_alignment

# Write data rows with alternating fills and center alignment
for row_idx, row in enumerate(final_df.itertuples(index=False), start=2):
    fill = light_row_fill if row_idx % 2 == 0 else dark_row_fill
    for col_idx, value in enumerate(row, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.fill = fill
        cell.alignment = center_alignment

# Freeze header and apply filter
ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

# Auto-fit column widths
for col in ws.columns:
    max_length = 0
    column = col[0].column
    column_letter = get_column_letter(column)
    for cell in col:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[column_letter].width = max_length + 2

# Save final styled Excel file
wb.save(output_excel)
print(f"✅ Styled Excel file saved as:\n{output_excel}")
