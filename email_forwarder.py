import imaplib
import email
import csv
import smtplib
import os 

accounts_path = f"{os.path.dirname(os.path.realpath(__file__))}/email_accounts.csv"

with open(accounts_path) as f:
    account_list = list(csv.reader(f))

for index, item in enumerate(account_list):
    if index == 0:
        continue
    id, IMAP_server, email_id, pwd, port, mailbox, prev_id, forwardee = item

    ############### IMAP SSL ##############################
    with imaplib.IMAP4_SSL(host=IMAP_server, port=port) as imap_ssl:
        print("Connection Object : {}".format(imap_ssl))

        ############### Login to Mailbox ######################
        print("Logging into mailbox...")
        resp_code, response = imap_ssl.login(email_id, pwd)

        print("Response Code : {}".format(resp_code))
        print("Response      : {}\n".format(response[0].decode()))

        ############### Set Mailbox #############
        resp_code, mail_count = imap_ssl.select(mailbox=mailbox, readonly=True)

        ############### Retrieve Mail IDs for given Directory #############
        resp_code, mails = imap_ssl.search(None, "ALL")
        # print("Mail IDs : {}\n".format(mails[0].decode().split()))

        ############### Select unread email IDs #############
        mail_list = mails[0].decode().split()
        max_id = len(mail_list)
        last_id = int(prev_id)
        max_count = min(max_id - last_id, 5)

        if max_count <= 0:
            print("All emails updated, no email to forward.")
            print("\nClosing selected mailbox....")
            imap_ssl.close()
            continue
        else:
            unread_mails = mail_list[-max_count:]


        ############### Read each email #############
        for mail_id in unread_mails:
            print(f"================== Start of Mail [{mail_id}] ====================")
            resp_code, mail_data = imap_ssl.fetch(mail_id, '(RFC822)')  # Fetch mail data.
            message = email.message_from_bytes(mail_data[0][1])  # Construct Message from mail data
            _from = f"From\t\t\t: {message.get('From')}"
            _date = f"Date\t\t\t: {message.get('Date')}"
            _subject = f"Subject\t\t\t: {message.get('Subject')}"
            _body = f"Message : \n\n{message.get_payload()}"
            full_message = f"{_from}\n{_date}\n{_subject}\n{_body}"
            print(full_message)
            print(f"================== End of Mail [{mail_id}] ====================\n")
            with smtplib.SMTP("smtp.gmail.com") as connection:
                connection.starttls()
                connection.login("mage.automailer@gmail.com", "Magical2019")
                connection.sendmail(
                    from_addr=message.get('From'),
                    to_addrs=forwardee,
                    msg=f'Subject: Forwarded from: {email_id}\n\n{full_message}')

        ############# Close Selected Mailbox #######################
        print("\nClosing selected mailbox....")
        imap_ssl.close()
        ############# Update last read email ID and save to CSV #######################
        account_list[index][-2] = max_id


with open(accounts_path, 'w', newline='') as f:
    print("Saving new account list...")
    csv_writer = csv.writer(f)
    csv_writer.writerows(account_list)
    print("Account list updated")
