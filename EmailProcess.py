#!/usr/bin/python

# TODO: Add gmail functionality.

import imapclient
import pyzmail
import pprint
import smtplib
import sys

class EmailProcess:
    imap_client = imapclient.IMAPClient("imap.mail.yahoo.com", ssl=True)
    uids = []
    user_email = ""
    password = ""

    def run(self):
        cls.login()
        cls.select_folder()
        cls.search()
        cls.sub_menu()

    def print_nnl(cls, input):
        # print no new line
        sys.stdout.write(input)
        sys.stdout.flush()

    def login(cls):
        while True:
            cls.user_email = raw_input("Please enter your yahoo email: ")
            cls.password = raw_input("Please enter your yahoo password: ")
            try:
                cls.print_nnl("Attepmting to login...")
                result = cls.imap_client.login(cls.user_email, cls.password)
            except cls.imap_client.Error, e:
                print("\nLogin error due to the following reasons:")
                print str(e)
                print "Please try again.\n"
                continue
            else:
                print "DONE."
                print result
                break

    def select_folder(cls):
        while True:
            print "Listing folders..."
            pprint.pprint(cls.imap_client.list_folders())
            folder_selection = raw_input("\nPlease enter one of the folders listed: ")
            try:
                cls.print_nnl("Selecting " + folder_selection + "...")
                cls.imap_client.select_folder(folder_selection, readonly=False)
            except cls.imap_client.Error, e:
                print "FAILED."
                print "\nError selecting folder due to the following reasons:"
                print str(e)
                print "Please try again.\n"
                continue
            else:
                print "DONE."
                print folder_selection, "selected.\n"
                break

    def search(cls):
        print("Please enter a search flag. Valid options include ALL, BEFORE, "
        "SINCE, SUBJECT, BODY, TEXT, FROM, CC, BCC, SEEN, UNSEEN, ANSWERED, "
        "UNANSWERED, DELETED, UNDELETED, DRAFT, UNDRAFT, FLAGGED, UNFLAGGED, "
        "LARGER, SMALLER NOT, and OR. Format for DATE is DD-MMM-YYYY. "
        "MMM is the first three letters of the month.")
        search_flag = raw_input("Flag: ")
        if search_flag == 'ALL':
            try:
                cls.print_nnl("Searching for all emails...")
                cls.uids = cls.imap_client.search([search_flag])
            except cls.imap_client.Error, e:
                print "\nError searching for all emails due to the following "
                "reason(s):"
                print str(e)
                print "Please try again.\n"
            else:
                print "Done."
                print "Found", len(cls.uids), "emails."
        else:
            search_term = raw_input("Please enter a search term: ")
            try:
                cls.print_nnl("Searching" + search_flag + search_term + "...")
                cls.uids = cls.imap_client.search([search_flag, search_term])
            except cls.imap_client.Error, e:
                print "\nError searching for emails" + search_flag + search_term
                "due to the following reason(s):"
                print e(str)
                print "Please try again.\n"
            else:
                print "Done."
                print "Found", len(cls.uids), "emails matching", search_term
                return cls.uids

    def sub_menu(cls):
        while True:
            print "1. Show email addresses and subjects."
            print "2. Attempt to send an unsubscribe message."
            print "3. Delete these emails."
            print "4. Main menu."
            choice = raw_input("Enter your chioce [1-3]: ")
            if choice == "1":
                cls.get_email_info()
            elif choice == "2":
                cls.send_unsubscribe()
            elif choice == "3":
                cls.delete_emails()
            elif choice == "4":
                return
            else:
                print "\n" + choice + " is an invalid selection. Please try again."
                continue

    def send_unsubscribe(cls):
        using_tls = False
        try:
            cls.print_nnl("Trying creation of smtp object for smtp.mail.yahoo.com "\
            "on port 587 with TLS...")
            smtp_object = smtplib.SMTP('smtp.mail.yahoo.com', 587)
        except smtplib.SMTPException, e:
            print "Error trying smtp.mail.yahoo.com on port 587 using TLS:"
            print str(e)
            try:
                cls.print_nnl("Trying creating of smtp object for"\
                "smtp.mail.yahoo.com on port 465 with SSL...")
                smtp_object = smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465)
            except smtplib.SMTPException, e:
                print "Error trying smtp.mail.yahoo.com on port 465 using SSL:"
                print str(e)
            else:
                print "Done."
                using_tls = False
                print "SSL on 465 successful."
        else:
            print "Done."
            using_tls = True
            print "TLS on 587 successful."
        try:
            cls.print_nnl("Saying hello...")
            result = smtp_object.ehlo()
        except smtplib.SMTPException, e:
            print "EHLO error:"
            print str(e)
        else:
            print "Done."
            print "EHLO successful."
            print result
        if using_tls:
            print "Starting TLS..."
            result = smtp_object.starttls()
            print result
        while True:
            cls.user_email = raw_input("Please enter your yahoo email: ")
            cls.password = raw_input("Please enter your yahoo password: ")
            try:
                cls.print_nnl("Attempting SMTP login...")
                result = smtp_object.login(cls.user_email, cls.password)
            except smtplib.SMTPException, e:
                print "\nError logging into SMTP server:"
                print str(e)
                print "Please try again.\n"
                continue
            else:
                print "Successful.\n"
                print result
                break
        print "Getting recipients..."
        emails = cls.get_recipients()
        print "Got these:"
        print emails
        for recipient in emails:
            print "Sending unsubscribe message to " + recipient + "..."
            result = smtp_object.sendmail(cls.user_email, recipient, 'Subject: UNSUBSCRIBE\n')
            print result
        result = smtp_object.quit()
        print result
        cls.sub_menu()

    def delete_emails(cls):
        try:
            cls.print_nnl("Attempting to flag", len(cls.uids),\
            "messages for deletion...")
            cls.imap_client.delete_messages(cls.uids)
        except cls.imap_client.Error, e:
            print "Error flagging messages:"
            print str(e)
        else:
            print "Done."
            while True:
                choice = raw_input("Messages ready for permanent deletion. Are you sure? Y/N")
                if choice in ("Y", "y"):
                    cls.print_nnl("Permanently deleting messages...")
                    result = cls.imap_client.expunge()
                    print "Done."
                    print result
                    break
                elif choice in ("N", "n"):
                    cls.sub_menu()
                else:
                    print "\n" + choice + " is an invalid selection. Please try again."
                    continue
        try:
            print "Permanently deleting emails..."
            cls.imap_client.expunge()
        except cls.imap_client.Error:
            print "Error deleting messages."
        else:
            print "Done."

    def get_email_info(cls):
        addresses = []
        subjects = []
        rawMessages = cls.imap_client.fetch(cls.uids, ['BODY[]'])
        for uid_key in cls.uids:
            message = pyzmail.PyzMessage.factory(rawMessages[uid_key]['BODY[]'])
            addresses.append(message.get_addresses('from'))
            subjects.append(message.get_subject())
        email_list = remove_duplicates(addresses)
        for subject in subjects:
            print subject
        print "Found", len(email_list), " email addresses."
        for address in email_list:
            print address

    def get_recipients(cls):
        # I think uids is coming up empty here. Fix it.
        emails = []
        rawMessages = cls.imap_client.fetch(cls.uids, ['BODY[]'])
        for number in cls.uids:
            message = pyzmail.PyzMessage.factory(rawMessages[number]['BODY[]'])
            emails.append(message.get_addresses('from'))
        emails = remove_duplicates(emails)
        return emails

    def remove_duplicates(cls, duplicates):
        email_list = []
        for address in duplicates:
            if address not in email_list:
                email_list.append(address)
        return email_list
