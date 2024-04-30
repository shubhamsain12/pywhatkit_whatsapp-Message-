

import openpyxl
import pywhatkit as kit

# Load the Excel file and select the active sheet
excel_file = r"C:\Users\asus\Desktop\numbers\numbers1.xlsx"
workbook = openpyxl.load_workbook(excel_file)
sheet = workbook.active

# Iterate through each row in the Excel sheet starting from row 2 (skipping header)
for row in sheet.iter_rows(min_row=2, max_col=2, values_only=True):
    raw_phone_number = str(row[0])  # Assuming phone numbers are in the first column
    message = '''Dosti nibhana ana cheye '''  # Your message here

    # Check if the phone number starts with a plus sign, if not, add it
    phone_number = raw_phone_number if raw_phone_number.startswith("+") else "+" + raw_phone_number

    # Ensure message is not empty
    if message.strip():

        try:
            # Send the message with pywhatkit instantly
            kit.sendwhatmsg_instantly(phone_number, message)
            print(f"Message sent successfully to {phone_number}")
        except Exception as e:
            print(f"Error sending message to {phone_number}: {e}")

# Close the Excel file
workbook.close()                

