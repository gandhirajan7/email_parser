import imaplib
import email
import datetime


def print_hi(name):
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


if __name__ == '__main__':
    print_hi('PyCharm')

    #login to oulook
    username = "gn479247@dal.ca"
    password = "Ganye@1997"
    myMail = imaplib.IMAP4_SSL('outlook.office365.com')
    myMail.login(username, password)
    myMail.select("INBOX") #select mailbox

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

    #get today's date
    date_today = datetime.datetime.now()
    date = date_today.strftime('%d-%b-%Y')
    # print(date)
    # resp, msgs = myMail.search(None, 'FROM "Gandhi Rajan"')
    # print(f'{resp} -- {msgs}')


    #fetch mail by sender and date
    # myMail.select("Inbox")
    # _, msgs = myMail.search(None,'ALL')
    # print(_,msgs)
    resp, msgs = myMail.search(None, f'ON {date} FROM "Gandhi Rajan"')
    for msg_no in msgs[0].split():
        _, msg = myMail.fetch(msg_no, '(RFC822)')
        # print(msg)
        message = email.message_from_bytes(msg[0][1])  # Parse the raw email message in to a convenient object
        print('\n')
        # print(message["subject"])
        subject = message["subject"].split("|")
        print(f'{len(subject)} -- {subject}')
        # sender = message["from"].split()[-1]

        #Email Details
        # print("== Email Details ====")
        # print(f'From: {message["from"]}')
        # print(f'sender {sender}\n')
        # print(f'Object type: {type(message)}')
        # print(f'Content type: {message.get_content_type()}')
        # print(f'Content disposition: {message.get_content_disposition()}')
        # print(f'Multipart?: {message.is_multipart()}')
        # if message.is_multipart():
        #     print("Types")
        #     for part in message.walk():
        #         print(f'- {part.get_content_type()}')
        #     multipart_payload = message.get_payload()
        #     for sub_msg in multipart_payload:
        #         pass
        #         # print(f'Payload\n{sub_msg.get_payload()}')
        # else:
        #     pass
        #     # print(f'Payload\n {message.get_content_type()}')




