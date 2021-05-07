import imaplib
import email
from email.header import decode_header
import webbrowser
import os


def print_hi(name):
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


if __name__ == '__main__':
    print_hi('PyCharm')
    username = "gandhirajan1997@gmail.com"
    password = "Gandhi@1234"
    myMail = imaplib.IMAP4_SSL('outlook.office365.com')
    myMail.login(username, password)
    myMail.select("INBOX")

    #print all messages
    # _, messages = myMail.search(None, 'ALL')
    # for message in messages[0].split():
    #     _, msg = myMail.fetch(message, '(RFC822)')
    #     print(msg[0][1])

    #list all folders
    response, folders = myMail.list()
    print(response)
    print("Available Folders")
    for folder in folders:
        # print(folder)
        folder_details = folder.decode().split()
        print(f'- {folder_details[-1]}')


