import csv
import openpyxl

# Example list of CSV-compatible lines
csv_data = [
    '"Name";"Age";"Location"',
    '"John";"25";"New York"',
    '"Alice";"30";"Los Angeles";"testabc";"testdef"',
    '"Bob";"28";"Chicago"'
]

# Create a new Excel workbook
wb = openpyxl.Workbook()
ws = wb.active
for row, line in enumerate(csv_data, start=1):
    for col, v2 in enumerate(csv.reader([line], delimiter=";", quotechar='"').__next__(), start=1):
        ws.cell(row, col, v2)

wb.save("test.xlsx")