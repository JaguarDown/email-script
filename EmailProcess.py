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
        self.login()
        self.select_folder()
        self.search()
        self.sub_menu()



    def print_nnl(self, input):
        # print no new line
        sys.stdout.write(input)
        sys.stdout.flush()

    def login(self):
        while True:
            self.user_email = raw_input("Please enter your yahoo email: ")
            self.password = raw_input("Please enter your yahoo password: ")
            try:
                self.print_nnl("Attepmting to login...")
                result = self.imap_client.login(self.user_email, self.password)
            except self.imap_client.Error, e:
                print("\nLogin error due to the following reasons:")
                print str(e)
                print "Please try again.\n"
                continue
            else:
                print "DONE."
                print result
                break

    def select_folder(self):
        while True:
            print "Listing folders..."
            pprint.pprint(self.imap_client.list_folders())
            folder_selection = raw_input("\nPlease enter one of the folders listed: ")
            try:
                self.print_nnl("Selecting " + folder_selection + "...")
                self.imap_client.select_folder(folder_selection, readonly=False)
            except self.imap_client.Error, e:
                print "FAILED."
                print "\nError selecting folder due to the following reasons:"
                print str(e)
                print "Please try again.\n"
                continue
            else:
                print "DONE."
                print folder_selection, "selected.\n"
                break

    def search(self):
        print("Please enter a search flag. Valid options include ALL, BEFORE, "
        "SINCE, SUBJECT, BODY, TEXT, FROM, CC, BCC, SEEN, UNSEEN, ANSWERED, "
        "UNANSWERED, DELETED, UNDELETED, DRAFT, UNDRAFT, FLAGGED, UNFLAGGED, "
        "LARGER, SMALLER NOT, and OR. Format for DATE is DD-MMM-YYYY. "
        "MMM is the first three letters of the month.")
        search_flag = raw_input("Flag: ")
        if search_flag == 'ALL':
            try:
                self.print_nnl("Searching for all emails...")
                self.uids = self.imap_client.search([search_flag])
            except self.imap_client.Error, e:
                print "\nError searching for all emails due to the following "
                "reason(s):"
                print str(e)
                print "Please try again.\n"
            else:
                print "Done."
                print "Found", len(self.uids), "emails."
        else:
            search_term = raw_input("Please enter a search term: ")
            try:
                self.print_nnl("Searching" + search_flag + search_term + "...")
                self.uids = self.imap_client.search([search_flag, search_term])
            except self.imap_client.Error, e:
                print "\nError searching for emails" + search_flag + search_term
                "due to the following reason(s):"
                print e(str)
                print "Please try again.\n"
            else:
                print "Done."
                print "Found", len(self.uids), "emails matching", search_term
                return self.uids

    def sub_menu(self):
        while True:
            print "1. Show email addresses and subjects."
            print "2. Attempt to send an unsubscribe message."
            print "3. Delete these emails."
            print "4. Main menu."
            choice = raw_input("Enter your chioce [1-3]: ")
            if choice == "1":
                self.get_email_info()
            elif choice == "2":
                self.send_unsubscribe()
            elif choice == "3":
                self.delete_emails()
            elif choice == "4":
                return
            else:
                print "\n" + choice + " is an invalid selection. Please try again."
                continue

    def send_unsubscribe(self):
        using_tls = False
        try:
            self.print_nnl("Trying creation of smtp object for smtp.mail.yahoo.com "\
            "on port 587 with TLS...")
            smtp_object = smtplib.SMTP('smtp.mail.yahoo.com', 587)
        except smtplib.SMTPException, e:
            print "Error trying smtp.mail.yahoo.com on port 587 using TLS:"
            print str(e)
            try:
                self.print_nnl("Trying creating of smtp object for"\
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
            self.print_nnl("Saying hello...")
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
            self.user_email = raw_input("Please enter your yahoo email: ")
            self.password = raw_input("Please enter your yahoo password: ")
            try:
                self.print_nnl("Attempting SMTP login...")
                result = smtp_object.login(self.user_email, self.password)
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
        emails = self.get_recipients()
        print "Got these:"
        print emails
        for recipient in emails:
            print "Sending unsubscribe message to " + recipient + "..."
            result = smtp_object.sendmail(self.user_email, recipient, 'Subject: UNSUBSCRIBE\n')
            print result
        result = smtp_object.quit()
        print result
        self.sub_menu()

    def delete_emails(self):
        try:
            self.print_nnl("Attempting to flag", len(self.uids),\
            "messages for deletion...")
            self.imap_client.delete_messages(self.uids)
        except self.imap_client.Error, e:
            print "Error flagging messages:"
            print str(e)
        else:
            print "Done."
            while True:
                choice = raw_input("Messages ready for permanent deletion. Are you sure? Y/N")
                if choice in ("Y", "y"):
                    self.print_nnl("Permanently deleting messages...")
                    result = self.imap_client.expunge()
                    print "Done."
                    print result
                    break
                elif choice in ("N", "n"):
                    self.sub_menu()
                else:
                    print "\n" + choice + " is an invalid selection. Please try again."
                    continue
        try:
            print "Permanently deleting emails..."
            self.imap_client.expunge()
        except self.imap_client.Error:
            print "Error deleting messages."
        else:
            print "Done."

    def get_email_info(self):
        addresses = []
        subjects = []
        rawMessages = self.imap_client.fetch(self.uids, ['BODY[]'])
        for uid_key in self.uids:
            message = pyzmail.PyzMessage.factory(rawMessages[uid_key]['BODY[]'])
            addresses.append(message.get_addresses('from'))
            subjects.append(message.get_subject())
        email_list = remove_duplicates(addresses)
        for subject in subjects:
            print subject
        print "Found", len(email_list), " email addresses."
        for address in email_list:
            print address

    def get_recipients(self):
        # I think uids is coming up empty here. Fix it.
        emails = []
        rawMessages = self.imap_client.fetch(self.uids, ['BODY[]'])
        for number in self.uids:
            message = pyzmail.PyzMessage.factory(rawMessages[number]['BODY[]'])
            emails.append(message.get_addresses('from'))
        emails = remove_duplicates(emails)
        return emails

    def remove_duplicates(self, duplicates):
        email_list = []
        for address in duplicates:
            if address not in email_list:
                email_list.append(address)
        return email_list
