#!/usr/bin/python

from EmailProcess import EmailProcess

def print_menu():
    print 30 * "-", "MENU", 30 * "-"
    print "1. Login to an email account"
    print "2. Quit"
    print 67 * "-"

def main():
    while True:
        print_menu()
        choice = raw_input("Enter your chioce [1-2]: ")
        if choice == "1":
            email_process = EmailProcess()
            email_process.run()
        elif choice == "2":
            quit()
        else:
            print "\n" + choice + " is an invalid selection. Please try again."
            continue

if __name__ == '__main__':
    main()
