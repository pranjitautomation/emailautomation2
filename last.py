import glob
import os
import smtplib
import ssl
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import geopy.distance
import pandas as pd
import schedule
from fpdf import FPDF
from geopy.geocoders import Nominatim
from imap_tools import AND, MailBox

# select the subject of email
raw_subject = "Distance between cities"
email_id = "pranjit.automation@gmail.com"
password = "Bubai97@"
email_server = "imap.gmail.com"


def login_to_email(email_server, email_id, password):
    # Server is the address of the imap server
    mb = MailBox(email_server).login(email_id, password)

    messages = mb.fetch(criteria=AND(seen=False),
                        mark_seen=True,
                        bulk=True)
    print("done")
    return messages


def data_extract_from_excel():
    for file in glob.glob("*.xlsx"):
        df = pd.read_excel(file)

        print(df["Cities Name"])

        city = list(set(df["Cities Name"]))
        print(city)
        os.remove(file)
        return city


def distance_between_cities():

    """
    This method extracts the city name from Excel
    And finds out the distance between two cities


    Returns:
        list : list containing distances between cities
    """

    city = data_extract_from_excel()
    # Initialize Nominatim API
    geolocator = Nominatim(user_agent="MyApp")

    distance_list = []

    for x in range(0, len(city)-1):

        location = geolocator.geocode(city[x])

        # finding the coordinate of location1
        coords_1 = (location.latitude, location.longitude)

        for z in range(x+1, len(city)):

            # finding the coordinate of location2
            location2 = geolocator.geocode(city[z])
            coords_2 = (location2.latitude, location2.longitude)
            find_dist = geopy.distance.distance(coords_1, coords_2).km

            final = f'distance between {city[x]} & {city[z]} is:-{find_dist}'
            distance_list.append(final)

    print(distance_list)

    return distance_list


def putting_data_of_distance_in_pdf():

    distance_list = distance_between_cities()
    pdf = FPDF()

    # Add a page
    pdf.add_page()

    for d in range(0, len(distance_list)):

        # set style and size of font
        pdf.set_font("Arial", size=15)

        # create a cell
        pdf.cell(200, 10, txt=distance_list[d],
                 ln=d+1, align='C')

    # save the pdf with name .pdf
    pdf.output("Distance_between_cities.pdf")


def send_pdf_in_mail(msg):
    """
    This method is used to send attachments in mail
    to the sender of excel file

    """

    subject = "Distance_Report_of_cities456"
    body = "Hi. This is the pdf Report of excel data"

    message = MIMEMultipart()

    message["From"] = msg.to[0]
    # Sending from that email_id in which excel file was received

    message["To"] = ", ".join(msg.from_)
    # sending to that email_id from which excel file was sent

    message["Subject"] = subject
    message.attach(MIMEText(body, "plain"))

    with open("Distance_between_cities.pdf", "rb") as attachment:
        attachment_part = MIMEBase("application", "octet-stream")
        attachment_part.set_payload(attachment.read())
        encoders.encode_base64(attachment_part)
        attachment_part.add_header(
            "Content-Disposition",
            "attachment; filename = Distance_between_cities.pdf",
        )
    message.attach(attachment_part)
    text = message.as_string()

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(msg.to[0], password)
        server.sendmail(msg.to[0], msg.from_, text)

    os.remove("Distance_between_cities.pdf")


def download_excel_and_sending_pdf():

    """
    This method downloads the excel file
    from that email id which contains the raw subject

    And Send the pdf to the same mail id

    """

    messages = login_to_email(email_server, email_id, password)
    for msg in messages:
        if msg.subject == raw_subject:

            print(msg.from_, ': ', msg.subject)
            print(msg.to)

            for att in msg.attachments:
                # downloading the excel file
                with open("./" + format(att.filename), 'wb') as f:
                    f.write(att.payload)

                # running each and every function for every mail
                putting_data_of_distance_in_pdf()

                send_pdf_in_mail(msg)

    print("OKKKK")


def func():
    # Scheduling code after a particuler time period
    download_excel_and_sending_pdf()


schedule.every(10).minutes.do(func)

while True:
    # It checks if any run is missed
    schedule.run_pending()
    time.sleep(1)
