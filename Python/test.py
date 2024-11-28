import smtplib
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Function to add text to a PDF and create a new PDF
def create_certificate(template_path, output_path, name, event_name):
    # Temporary PDF to add text
    temp_pdf_path = "temp.pdf"

    # Create a new PDF with the text overlay
    c = canvas.Canvas(temp_pdf_path, pagesize=letter)
    c.setFont("Helvetica", 20)  # Adjust font size as needed
    c.setFillColorRGB(0, 0, 0)  # Set text color to black

    # Adjust coordinates for the text placement
    c.drawString(360, 300, f" {name}")  # Example position for the name
    c.drawString(360, 255, f"{event_name}")  # Example position for the event name
    c.save()

    # Read the template and overlay the new text
    template = PdfReader(template_path)
    temp_pdf = PdfReader(temp_pdf_path)
    writer = PdfWriter()

    # Copy the template page
    template_page = template.pages[0]
    template_page.merge_page(temp_pdf.pages[0])  # Merge text with template
    writer.add_page(template_page)

    # Save the final output
    with open(output_path, "wb") as output_pdf:
        writer.write(output_pdf)

    # Remove temporary file
    os.remove(temp_pdf_path)

# Function to send email
def send_email(sender_email, sender_password, recipient_email, subject, body, attachments):
    try:
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        # Attach files
        for file in attachments:
            with open(file, "rb") as f:
                mime_base = MIMEBase('application', 'octet-stream')
                mime_base.set_payload(f.read())
            encoders.encode_base64(mime_base)
            mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file)}')
            msg.attach(mime_base)

        # Connect and send email
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}. Error: {e}")

# Main function
def main():
    # Email credentials
    sender_email = "adroitcsd@gmail.com"
    sender_password = "rtfz odom btvz ipmt"  # Use your app password here

    # Paths
    template_pdf = "D:/Python/certificate_template.pdf"  # Replace with your template path
    excel_file = "D:/Python/Certificate.xlsx"  # Replace with your Excel file path

    # Read Excel file
    data = pd.read_excel(excel_file)

    for index, row in data.iterrows():
        recipient_email = row['Email']
        names = [row['Name1'], row['Name2'], row['Name3'], row['Name4']]
        event_name = row['EventName']

        # Create certificates
        certificates = []
        for name in names:
            output_pdf = f"{name}_certificate.pdf"
            create_certificate(template_pdf, output_pdf, name, event_name)
            certificates.append(output_pdf)

        # Email subject and body
        subject = f"Certificates for {event_name}"
        body = f"Dear Participant,\n\nPlease find attached the certificates for {', '.join(names)} from the {event_name} event.\n\nBest Regards,\nEvent Team"

        # Send email
        send_email(sender_email, sender_password, recipient_email, subject, body, certificates)

        # Remove generated certificates
        for cert in certificates:
            os.remove(cert)

if __name__ == "__main__":
    main()



# import smtplib
# import pandas as pd
# from PIL import Image, ImageDraw, ImageFont
# from email.mime.multipart import MIMEMultipart
# from email.mime.text import MIMEText
# from email.mime.base import MIMEBase
# from email import encoders
# import os
#
#
# # Function to create a certificate image
# def create_certificate(template_path, output_path, name, event_name, font_path="arial.ttf"):
#     # Open the image template
#     img = Image.open(template_path)
#     draw = ImageDraw.Draw(img)
#
#     # Define font and text positions
#     font_name = ImageFont.truetype(font_path, 40)  # Adjust font size
#     font_event = ImageFont.truetype(font_path, 30)
#
#     # Coordinates for the name and event name (adjust these as needed)
#     name_coords = (250, 300)  # Replace with actual coordinates for name
#     event_coords = (250, 400)  # Replace with actual coordinates for event
#
#     # Add text to the image
#     draw.text(name_coords, name, fill="black", font=font_name)
#     draw.text(event_coords, f"Event: {event_name}", fill="black", font=font_event)
#
#     # Save the edited image
#     img.save(output_path)
#
#
# # Function to send email
# def send_email(sender_email, sender_password, recipient_email, subject, body, attachments):
#     try:
#         # Create the email
#         msg = MIMEMultipart()
#         msg['From'] = sender_email
#         msg['To'] = recipient_email
#         msg['Subject'] = subject
#         msg.attach(MIMEText(body, 'plain'))
#
#         # Attach files
#         for file in attachments:
#             with open(file, "rb") as f:
#                 mime_base = MIMEBase('application', 'octet-stream')
#                 mime_base.set_payload(f.read())
#             encoders.encode_base64(mime_base)
#             mime_base.add_header('Content-Disposition', f'attachment; filename={os.path.basename(file)}')
#             msg.attach(mime_base)
#
#         # Connect and send email
#         with smtplib.SMTP('smtp.gmail.com', 587) as server:
#             server.starttls()
#             server.login(sender_email, sender_password)
#             server.send_message(msg)
#             print(f"Email sent to {recipient_email}")
#     except Exception as e:
#         print(f"Failed to send email to {recipient_email}. Error: {e}")
#
#
# # Main function
# def main():
#     # Email credentials
#     sender_email = "adroitcsd@gmail.com"
#     sender_password = "rtfz odom btvz ipmt"
#
#     # Paths
#     template_image = "D:/Python/certificate_template.png"  # Replace with your template image path
#     excel_file = "D:/Python/Certificate.xlsx"  # Replace with your Excel file path
#
#     # Read Excel file
#     data = pd.read_excel(excel_file)
#
#     for index, row in data.iterrows():
#         recipient_email = row['Email']
#         names = [row['Name1'], row['Name2'], row['Name3'], row['Name4']]
#         event_name = row['EventName']
#
#         # Create certificates
#         certificates = []
#         for name in names:
#             output_image = f"{name}_certificate.jpg"
#             create_certificate(template_image, output_image, name, event_name)
#             certificates.append(output_image)
#
#         # Email subject and body
#         subject = f"Certificates for {event_name}"
#         body = f"Dear Participant,\n\nPlease find attached the certificates for {', '.join(names)} from the {event_name} event.\n\nBest Regards,\nEvent Team"
#
#         # Send email
#         send_email(sender_email, sender_password, recipient_email, subject, body, certificates)
#
#         # Remove generated certificates
#         for cert in certificates:
#             os.remove(cert)
#
#
# if __name__ == "__main__":
#     main()

