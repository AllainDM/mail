import imaplib
import email
# from email.header import decode_header
# import base64
# import re
# from datetime import datetime


# from bs4 import BeautifulSoup
# import xlrd
# import xlrd3
# import openpyxl
# import xlwt

import config
import userside
# import for_api
# import filter


# Настройка imaplib
mail_pass = config.password
username = config.address
imap_server = "imap.mail.ru"
imap = imaplib.IMAP4_SSL(imap_server)
imap.login(username, mail_pass)

# imap.select("INBOX")
imap.select('user')
# p = imap.search(None, 'ALL')
# r = imap.uid('search', "UNSEEN", "ALL")
typ, data = imap.uid('search', "UNSEEN", "ALL")


# Чтение почты и сохранение вложенных файлов в папку
def start():
    list_filenames = []  # Список имен полученных файлов
    for num in data[0].split():
        res, msg = imap.uid('fetch', num, '(RFC822)')

        mail = email.message_from_bytes(msg[0][1])

        if mail.is_multipart():
            for part in mail.walk():
                # content_type = part.get_content_type()
                filename = part.get_filename()
                # Найдем файлы с названием(строка)

                # if type(filename) == str and filename[16] == 'c':  # Первый вариант для userside
                if type(filename) == str:
                    print(filename)
                    print(type(filename))
                    list_filenames.append(filename)
                    with open(f'files/{filename}', 'wb') as new_file:
                        new_file.write(part.get_payload(decode=True))

    print(f"Получены файлы: {list_filenames}")

    # Запуск для userside по всем ТО без даты
    userside.econtracts("", "north", list_filenames)
    userside.econtracts("", "south", list_filenames)
    userside.econtracts("", "west", list_filenames)
    userside.econtracts("", "east", list_filenames)


if __name__ == '__main__':
    start()
