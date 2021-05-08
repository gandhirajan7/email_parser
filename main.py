import imaplib
import email
import datetime
from email.header import decode_header
import webbrowser
import os


def print_hi(name):
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


if __name__ == '__main__':
    print_hi('PyCharm')
    username = "gn479247@dal.ca"
    password = "Ganye@1997"
    myMail = imaplib.IMAP4_SSL('outlook.office365.com')
    myMail.login(username, password)
    myMail.select("INBOX")

    #print all messages
    # _, messages = myMail.search(None, 'ALL')
    # for message in messages[0].split():
    #     _, msg = myMail.fetch(message, '(RFC822)')
    #     print(msg[0][1])

    #list all folders
    # response, folders = myMail.list()
    # print("Available Folders")
    # for folder in folders:
    #     # print(folder)
    #     folder_details = folder.decode().split()
    #     print(f'- {folder_details[-1]}')

    #fetch a folder by sender
    # myMail.select("Inbox")
    # _, msgs = myMail.search(None,'ALL')
    # print(_,msgs)
    # for msg_no in msgs[0].split():
    #     _, msg = myMail.fetch(msg_no, '(RFC822)')
    #     # print(msg)
    #     message = email.message_from_bytes(msg[0][1])
    #     print(message["from"])
    #     sender = message["from"].split()[-1]
    #     if sender == "<gandhirajan1997@gmail.com>":
    #         print("== Email Details ====")
    #         # print(f'From: {message["from"]}')
    #         print(f'sender {sender}')

    # time = time.time()
    date_today = datetime.datetime.now()
    date = date_today.strftime('%d-%m-%Y')
    print(date)
    # resp, data = myMail.search(None, 'FROM "Gandhi Rajan <gandhirajan1997@gmail.com>"')
    resp, data = myMail.search(None, 'SUBJECT ""')
    print(data)


