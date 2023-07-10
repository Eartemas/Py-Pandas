# Please note that the code above assumes you have already installed the pandas and pywin32 libraries in your Python environment. 
# Additionally, make sure to replace the file names, email addresses, and other placeholders with the appropriate values based on your project's requirements.
# Before running the code, ensure that you have the required permissions to access the data source, create the Excel file, and send emails using Outlook.

import pandas as pd
import win32com.client

# Step a: Import necessary libraries

# Step b: Read data from Excel file
excel_file = 'sales_data.xlsx'  # Replace 'sales_data.xlsx' with your Excel file name
sheet_name = 'Sheet1'  # Replace 'Sheet1' with the appropriate sheet name in your Excel file

data = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

# Step c: Find the current week's data
current_week_data = []
start_row = 0
end_row = 0

for row in range(len(data)):
    if data.iloc[row].str.startswith("Week").any():
        start_row = row + 2  # Skip two rows for header and empty row
    elif data.iloc[row].str.contains("WEEKLY TOTAL").any():
        end_row = row + 1  # Include the row with 'WEEKLY TOTAL'
        break

if start_row != 0 and end_row != 0:
    current_week_data = data.iloc[start_row:end_row]

# Step d: Create Outlook email object
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0 represents a new mail item

# Step e: Set email properties
mail.Subject = "Weekly Sales Report"
mail.Body = "Please find the weekly sales report section attached."

# Set recipients (replace 'manager1@example.com', 'manager2@example.com', etc. with actual email addresses)
mail.To = "manager1@example.com; manager2@example.com"

# Step f: Export current week's data to Excel
current_week_excel_file = 'current_week_sales_data.xlsx'  # Replace with your desired Excel file name for the current week's data
current_week_sheet_name = 'Current Week Sales Data'  # Replace with your desired sheet name for the current week's data

current_week_data.to_excel(current_week_excel_file, sheet_name=current_week_sheet_name, index=False, header=False)

# Step g: Attach the Excel sheet to the email
attachment = current_week_excel_file
mail.Attachments.Add(attachment)

# Step h: Send the email
maIL.Send()
