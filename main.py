#!#!/usr/bin/env python

import os
import smtplib
import time
from email.headerregistry import Address
from email.message import EmailMessage
from email.mime.application import MIMEApplication
from argparse import ArgumentParser
from typing import Callable

import mammoth
import pandas as pd
from pandas import DataFrame

SELECTED_COLUMNS = ['DOCX', 'PDF_FILE', 'EMAIL', 'COMPANY']
FROM_ADDRESS = Address("William Pitts", "William", "chi-awe.org")

TO_ADDRESSES = (
                Address("William", "pittswilliam715", "gmail.com"),
                )


def get_smtp_client(username: str, password: str,
                    host: str, port: int) -> smtplib.SMTP:
    # credentials in here
    smtp = smtplib.SMTP(host, port)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(username, password)

    return smtp


def get_html_from_docx(filename: str) -> str:
    with open(filename, "rb") as docx_file:
        result = mammoth.convert_to_html(docx_file)
        html = result.value  # The generated HTML
        # messages = result.messages  # Any messages, such as warnings during conversion

        return html


def get_txt_from_docx(filename: str) -> str:
    with open(filename, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)

        return result.value


def attach(msg, filename):
    with open(filename, 'rb') as f:
        file = MIMEApplication(
            f.read(),
            name=os.path.basename(filename))
    msg.attach(file)


def generate_email_message(subject: str,
                           frm: Address,
                           to,
                           txt: str,
                           html: str) -> EmailMessage:
    msg = EmailMessage()

    msg['From'] = frm
    msg['To'] = to
    msg['Subject'] = subject

    msg.set_content(txt)
    msg.add_alternative(f"<html><head></head><body>{html}</body></html>", subtype='html')

    return msg


def to_address(display_name, email_address: str) -> Address:
    to_parts = email_address.strip().removesuffix(".").split("@")
    to = Address(display_name, to_parts[0], to_parts[1])

    return to


# send email => text + recipient + 2 attachments
def process(email_sender: Callable, df_selected: DataFrame, debug: bool = True):
    if debug:
        x = range(1)
    else:
        # read the start + 1
        offset = 0
        x = range(offset, len(df_selected))

    for i in x:
        if debug:
            to_addresses = TO_ADDRESSES
        else:
            to_addresses = (to_address(display_name=df_selected.COMPANY[i], email_address=df_selected.EMAIL[i]))

        # subject = f"MAC2024 Silent Auction Sponsorship: {df_selected.COMPANY[i]}"
        subject = f"Chicago Asian Women Empowerment"

        source_docx = df_selected.DOCX[i]
        source_docx = source_docx.replace("\\", os.sep)

        request_letter_file = df_selected.PDF_FILE[i]
        request_letter_file = request_letter_file.replace("\\", os.sep)

        msg = generate_email_message(subject=subject,
                                     frm=FROM_ADDRESS,
                                     to=to_addresses,
                                     txt=get_txt_from_docx(source_docx),
                                     html=get_html_from_docx(source_docx))
        attach(msg, filename=request_letter_file)
        attach(msg, filename="501c3 CAWE.pdf")

        email_sender(msg)
        # commit the state
        # print(i)
        time.sleep(1)


class EmailSender(Callable):
    def __init__(self, smtp: smtplib.SMTP, dry_run: bool = True):
        self.smtp = smtp
        self.dry_run = dry_run

    def __call__(self, msg: EmailMessage):
        if not self.dry_run:
            print(f"Sending message to: {msg['to']}")
            self.smtp.send_message(msg)
        else:
            print(f"Will send message to: {msg['to']}")


if __name__ == '__main__':
    parser = ArgumentParser()
    parser.add_argument("-c", "--csv", help="CSV file to parse", required=True)
    parser.add_argument("-s", "--send", help="Send Emails or not", required=False, default=False,
                        action='store_true')
    parser.add_argument("-a", "--actual", help="Use actual emails from CSV or hard coded test emails", required=False, default=False,
                        action='store_true')
    args = parser.parse_args()

    csv_file = args.csv
    dry_run = not args.send
    debug = not args.actual

    if not dry_run:
        # generate from env
        username = os.getenv("SMTP_USERNAME")
        password = os.getenv("SMTP_PASSWORD")
        host = os.getenv("SMTP_HOST")
        port = int(os.getenv("SMTP_PORT"))
        # login
        smtp = get_smtp_client(username, password, host, port)
    else:
        smtp = None

    # Load the dataset
    df = pd.read_csv(csv_file)
    df_selected = df[SELECTED_COLUMNS]

    if debug:
        # Display the first few rows of the new DataFrame to verify / debug
        print(df_selected.head())

    process(EmailSender(smtp, dry_run), df_selected, debug)

    if smtp:
        smtp.quit()
