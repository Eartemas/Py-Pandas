# Please note that the code above assumes you have already installed the pandas and pywin32 libraries in your Python environment. 
# Additionally, make sure to replace the file names, email addresses, and other placeholders with the appropriate values based on your project's requirements.
# Before running the code, ensure that you have the required permissions to access the data source, create the Excel file, and send emails using Outlook.

import pandas as pd
import win32com.client   # use pip install pywin32 to install pywin32

# Step a: Import necessary libraries

# Step b: Read data from Excel file
excel_file = 'sales_data.xlsx'  # Replace 'sales_data.xlsx' with your Excel file name
sheet_name = 'Sheet1'  # Replace 'Sheet1' with the appropriate sheet name in your Excel file

data = pd.read_excel(excel_file, sheet_name=sheet_name)

# Step c: Copy the desired section of the data
start_row = 2  # Replace with the row number where your desired section starts
end_row = 10  # Replace with the row number where your desired section ends

copied_data = data.iloc[start_row - 1: end_row]  # Select the desired rows (adjusting for 0-based indexing)

# Step d: Create Outlook email object
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0 represents a new mail item

# Step e: Set email properties
mail.Subject = "Weekly Sales Report"
mail.Body = "Please find the weekly sales report section attached."

# Set recipients replace 'manager1@example.com', 'manager2@example.com', etc. 
# with actual email addresses
mail.To = "manager1@example.com; manager2@example.com"

# Step f: Export copied data to Excel
copied_data_excel_file = 'copied_sales_data.xlsx'  # Replace with your desired Excel file name for the copied data
copied_data_sheet_name = 'Copied Sales Data'  # Replace with your desired sheet name for the copied data

copied_data.to_excel(copied_data_excel_file, sheet_name=copied_data_sheet_name, index=False)

# Step g: Attach the Excel sheet to the email
attachment = copied_data_excel_file
mail.Attachments.Add(attachment)

# Step h: Send the email
mail.Send()
