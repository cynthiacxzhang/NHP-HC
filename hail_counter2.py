#work up from the last row --> code is reading empty excel rows (not what we want)
#xlsx file limitations?? 
#working way up - only method currenlty without reading empty cells

import openpyxl

def check_recent_values(filename, sheetname, column_name, num_values):
    # Load the workbook and select the specified sheet
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook[sheetname]
    
    # Initialize flags to track positive and negative values
    positive_count = 0
    negative_value = False
    
    # Start from the last row and work our way up
    for row in range(sheet.max_row, 0, -1):
        cell_value = sheet[column_name + str(row)].value
        if cell_value is not None:
            if cell_value > 0:
                positive_count += 1
                if positive_count >= num_values:
                    return True
            elif cell_value < 0:
                negative_value = True
    
    if negative_value:
        return "error- station failed"
    
    return False

def hail_counter():
    print("There has been hail in the past hour.")

# Specify the filename, sheetname, and column name
filename = "data.xlsx" #new file with most recent data needed
sheetname = "Sheet"
column_name = "A"  # Replace with the actual column name

# Check if any of the most recent six values are greater than zero
result = check_recent_values(filename, sheetname, column_name, num_values=6)

if result == True:
    print("Found positive values, triggering hail_counter")
    hail_counter()
elif result == "error- station failed":
    print("Error - Station failed as negative value was detected")
else:
    print("No positive values found")
