#!/usr/bin/python

import imapclient
import pyzmail
import pprint
import smtplib
import sys
from ConsoleColor import ConsoleColor

class EmailScript:
    imap_client = imapclient.IMAPClient("imap.mail.yahoo.com", ssl=True)
    colors = ConsoleColor()
    is_logged_in = False
    uids = []
    user_email = ""
    password = ""

    def execute(self):
        self.colors.color_print("Welcome to JaguarDown's email script! Enter q"\
        " at any prompt to quit.")
        # LOGIN
        while True:
            if self.is_logged_in == False:
                self.login()

            # SELECT FOLDER
            while True:
                try:
                    print "Listing folders..."
                    pprint.pprint(self.imap_client.list_folders())
                except self.imap_client.Error:
                    self.colors.fail()
                    self.is_logged_in = False
                    break
                folder_selection = raw_input("Choose folder to search in: ")
                try:
                    self.print_nnl("Selecting '" + folder_selection + "'...")
                    self.imap_client.select_folder(folder_selection, readonly=False)
                except self.imap_client.Error, e:
                    self.colors.fail()
                    print "Couldn't select '" + folder_selection
                    print "Please try again."
                    continue
                else:
                    self.colors.ok()
                    break

            # SEARCH
            while True:
                search_flag = raw_input("Search flag (type 'help' to list options): ")
                if search_flag == 'help':
                    print("ALL, BEFORE, SINCE, SUBJECT, BODY, TEXT, FROM, CC, BCC, "\
                    "SEEN, UNSEEN, ANSWERED, UNANSWERED, DELETED, UNDELETED, DRAFT, "\
                    "UNDRAFT, FLAGGED, UNFLAGGED, LARGER, SMALLER NOT, and OR. Format "\
                    "for DATE is DD-MMM-YYYY. MMM is the first three letters of the "\
                    "month.")
                    search_flag = ""
                    continue
                elif search_flag == 'ALL':
                    try:
                        self.print_nnl("Searching for all emails...")
                        self.uids = self.imap_client.search([search_flag])
                    except self.imap_client.Error, e:
                        self.colors.fail()
                        print "Couldn't list all emails, please try again."
                        continue
                    else:
                        self.colors.ok()
                        print "Found", len(self.uids), "emails."
                else:
                    search_term = raw_input("Please enter a search term: ")
                    try:
                        self.print_nnl("Searching for emails '" + search_flag + "' '" +\
                        search_term + "'...")
                        self.uids = self.imap_client.search([search_flag, search_term])
                    except self.imap_client.Error, e:
                        self.colors.fail()
                        print "Error searching for emails" + search_flag + search_term + "."\
                        "Please try again."
                    else:
                        self.colors.ok()
                        print "Found", len(self.uids), "emails matching your search."
                        if len(self.uids) == 0:
                            continue
                        else:
                            break

            while True:
                choice = raw_input("Do you want to unsubscribe? [Y/n] ")
                if choice in ("Y", "y"):
                    self.unsubscribe()
                    break
                elif choice in ("N", "n"):
                    break
                else:
                    print "Huh?"

            while True:
                choice = raw_input("Do you want to delete " + str(len(self.uids)) + " emails? [Y/n] ")
                if choice in ("Y", "y"):
                    self.delete()
                    break
                elif choice in ("N", "n"):
                    break
                else:
                    print "Huh?"
                    continue

            while True:
                choice = raw_input("Would you like to start over? [Y/n] ")
                if choice in ("Y", "y"):
                    break
                elif choice in ("N", "n"):
                    quit()
                else:
                    print "Huh?"
                    continue

    def print_nnl(self, input):
        # print no new line
        sys.stdout.write(input)
        sys.stdout.flush()

    def login(self):
        while True:
            self.user_email = raw_input("Email: ")
            self.password = raw_input("Password: ")
            try:
                self.print_nnl("Logging in...")
                self.imap_client.login(self.user_email, self.password)
            except self.imap_client.Error:
                self.colors.fail()
                print "Invalid credentials. Please try again."
                continue
            else:
                self.is_logged_in = True
                self.colors.ok()
                break

    def unsubscribe(self):
        using_tls = False
        try:
            self.print_nnl("Connecting to smtp.mail.yahoo.com:587 with TLS...")
            smtp_object = smtplib.SMTP('smtp.mail.yahoo.com', 587)
        except smtplib.SMTPException:
            self.colors.fail()
            try:
                self.print_nnl("Connecting to smtp.mail.yahoo.com:465 with SSL...")
                smtp_object = smtplib.SMTP_SSL('smtp.mail.yahoo.com', 465)
            except smtplib.SMTPException:
                self.colors.fail()
                return
            else:
                self.colors.ok()
                using_tls = False
        else:
            self.colors.ok()
            using_tls = True
        try:
            self.print_nnl("Saying hello...")
            smtp_object.ehlo()
        except smtplib.SMTPException:
            self.colors.fail()
            return
        else:
            self.colors.ok()
        if using_tls:
            try:
                self.print_nnl("Starting TLS...")
                smtp_object.starttls()
            except smtplib.SMTPException:
                self.colors.fail()
                return
            else:
                self.colors.ok()
        try:
            self.print_nnl("Logging into the SMTP server with your credentials...")
            smtp_object.login(self.user_email, self.password)
        except smtplib.SMTPException:
            self.colors.fail()
            return
        else:
            self.colors.ok()
        self.print_nnl("Getting addresses from the emails...")
        emails = self.get_recipients()
        self.colors.ok()
        print "Found", len(emails), "email addresses."
        for recipient in emails:
            try:
                self.print_nnl("Sending unsubscribe message to " + recipient + "...")
                smtp_object.sendmail(self.user_email, recipient, 'Subject: UNSUBSCRIBE\n')
            except smtplib.SMTPServerDisconnected as e:
                self.colors.fail()
                smtp_object.connect()
                continue
            except smtplib.SMTPSenderRefused as e:
                self.colors.fail()
                continue
            else:
                self.colors.ok()

        smtp_object.quit()

    def get_recipients(self):
        emails = []
        rawMessages = self.imap_client.fetch(self.uids, ['BODY[]'])
        for number in self.uids:
            message = pyzmail.PyzMessage.factory(rawMessages[number]['BODY[]'])
            emails.append(message.get_addresses('from'))
        emails = self.remove_duplicates(emails)
        return emails

    def remove_duplicates(self, duplicates):
        email_addresses = []
        duplicate_addresses = []
        for sub_list in duplicates:
            email = [x[1] for x in sub_list]
            duplicate_addresses.append(email[0])
        for item in duplicate_addresses:
            if item not in email_addresses:
                email_addresses.append(item)
        return email_addresses

    def delete(self):
        try:
            self.print_nnl("Flagging " + str(len(self.uids)) + \
            " messages for deletion...")
            self.imap_client.delete_messages(self.uids)
        except self.imap_client.Error:
            self.colors.fail()
        else:
            self.colors.ok()
            while True:
                try:
                    self.print_nnl("Permanently deleting emails...")
                    self.imap_client.expunge()
                except self.imap_client.Error, e:
                    self.colors.fail()
                    break
                else:
                    self.colors.ok()
                    self.uids = []
                    break
