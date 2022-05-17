import imaplib
import email
import datetime
import xlsxwriter
from feedback import Feedback


def print_hi(name):
    print(f'Hi, {name}')  # Press Ctrl+F8 to toggle the breakpoint.


def create_connection(username, password):
    # login to oulook
    connection = imaplib.IMAP4_SSL('outlook.office365.com')
    connection.login(username, password)
    return connection


def print_all_mail(connection):
    # print all messages
    _, messages = connection.search(None, 'ALL')
    for mesg in messages[0].split():
        _, data = connection.fetch(mesg, '(RFC822)')
        print(data[0][1])


def get_mail_by_date(connection):
    # print(date)
    # resp, msgs = myMail.search(None, 'FROM "Gandhi Rajan"')
    # _, msgs = myMail.search(None,'ALL')
    # print(f'{resp} -- {msgs}')

    # fetch mail by sender and date
    connection.select("Inbox")
    resp, msgs = connection.search(None, f'ON {date} FROM "Gandhi Rajan"')
    feedback_arr = []
    for msg_no in msgs[0].split():
        _, msg = myMail.fetch(msg_no, '(RFC822)')
        message = email.message_from_bytes(msg[0][1])  # Parse the raw email message in to a convenient object
        # print('\n')
        subject = message["subject"].split("|")
        # print(f'{len(subject)} -- {subject}')
        # sender = message["from"].split()[-1]
        if len(subject) == 5:
            pId = subject[0]
            name = subject[1]
            mail = subject[2]
            product = subject[3]
            review = subject[4]
            feedback = Feedback(pId, name, mail, product, review)
            feedback_arr.append(feedback)
    return feedback_arr


def create_datasheet(feedback_list):
    file = xlsxwriter.Workbook(f"{date}.xlsx")
    sheet = file.add_worksheet(name="Feedback")
    row = column = 0
    headers = ["Product ID", "Customer Name", "Customer Mail ID", "Product Name", "Product ID"]
    for item in headers:
        sheet.write(row, column, item)
        column += 1
    row = 0
    for record in feedback_list:
        row += 1
        column = 0
        sheet.write(row, column, record.fNo)
        column += 1
        sheet.write(row, column, record.name)
        column += 1
        sheet.write(row, column, record.mail)
        column += 1
        sheet.write(row, column, record.product)
        column += 1
        sheet.write(row, column, record.review)

    file.close()


if __name__ == '__main__':
    print_hi('PyCharm')
    date_today = datetime.datetime.now() # get today's date
    date = date_today.strftime('%d-%b-%Y')
    myMail = create_connection("", "")
    feedback_list_for_the_day = get_mail_by_date(myMail)
    create_datasheet(feedback_list_for_the_day)

    # list all folders
    # response, folders = myMail.list()
    # print("Available Folders")
    # for folder in folders:
    #     # print(folder)
    #     folder_details = folder.decode().split()
    #     print(f'- {folder_details[-1]}')

    # Email Details
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
