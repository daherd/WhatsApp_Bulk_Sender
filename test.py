from openpyxl import load_workbook

# Load an existing workbook
workbook = load_workbook("Book1.xlsx")

# Get the active sheet
sheet = workbook.active

# Read data from a cell
print(sheet['A1'].value)
print(sheet['A2'].value)