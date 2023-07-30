import smtplib
import os
import chardet
import openpyxl

def send_email(email, password, subject, body, attachment):
    # Create a SMTP object
    smtp = smtplib.SMTP("smtp.live.com", 587)

    # Start TLS for security
    smtp.starttls()

    # Login to the SMTP server
    smtp.login(email, password)

    # Send the email
    message = f"From: {email}\nTo: {email}\nSubject: {subject}\n\n{body}"
    with open(attachment, "rb") as f:
        smtp.sendmail(email, email, message, f.read())

    # Close the SMTP connection
    smtp.quit()

def main():
    # Get the email and password from the user
    email = input("Enter your Outlook email: ")
    password = input("Enter your Outlook password or app password: ")

    # Get the path to the XLSX file
    xlsx_file_path = input("Enter the path to the XLSX file: ")

    # Get the path to the folder that contains the PDF files
    pdf_folder_path = input("Enter the path to the folder that contains the PDF files: ")

    # Open the XLSX file using openpyxl
    workbook = openpyxl.load_workbook(xlsx_file_path)

    # Select the first sheet in the workbook (you can adjust this if needed)
    sheet = workbook.active

    # Find the column index with the header "email"
    email_column_index = None
    for col_idx, cell in enumerate(sheet[1], start=1):  # Assuming the header is in the first row
        if cell.value == "email":
            email_column_index = col_idx
            break

    if email_column_index is None:
        print("Header 'email' not found in the XLSX file.")
        return

    # Get the list of emails from the identified column
    emails = [sheet.cell(row=row_idx, column=email_column_index).value for row_idx in range(2, sheet.max_row + 1)]

    # Loop through the emails
    for email in emails:
        # Get the name of the PDF file
        pdf_file_name = email.split("@")[0] + ".pdf"

        # Get the path to the PDF file
        pdf_file_path = os.path.join(pdf_folder_path, pdf_file_name)

        # Check if the PDF file exists
        if os.path.exists(pdf_file_path):
            # Send the email
            send_email(email, password, "Subject", "Body", pdf_file_path)
        else:
            print(f"Email {email} does not have an attachment and will not be sent.")

if __name__ == "__main__":
    main()
