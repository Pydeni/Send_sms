import email
import imaplib
from email.header import decode_header
import pandas as pd

import sys
import urllib
import requests


mail_pass = "wrerkbxcndbfxqni"
username = "sonic269@yandex.ru"
imap_server = "imap.yandex.ru"
imap = imaplib.IMAP4_SSL(imap_server)
imap.login(username, mail_pass)
# Входим в папку "Входящие"
imap.select("INBOX")
# Поиск непрочитанных писем
imap.search(None, "UNSEEN")
# Получаем номер последнего письма
nomer_last_mail = '0'
# Передаем номер письма, письмо становится прочитанным
res, msg = imap.fetch(b'8393', '(RFC822)')
# Извлекаем письмо
msg = email.message_from_bytes(msg[0][1])
# От кого, дата и время прихода получения письма, Айди
letter_from = msg["Return-path"]
letter_date = email.utils.parsedate_tz(msg["Date"])
letter_id = msg["Message-ID"]
# Декодируем тему письма
letter_tema = decode_header(msg["Subject"])[0][0].decode()
# Получаем инфу, что в письме
msg.get_payload()
msg.is_multipart()
"""Проходимся по вложению"""
# for part in msg.walk():
#     print(part.get_content_type())
# Имя вложения письма
name_vlojenie = ''
for part in msg.walk():
    if part.get_content_disposition() == 'attachment':
        name_vlojenie = decode_header(part.get_filename())[0][0].decode()
        filename = part.get_filename()
        if filename:
            with open(name_vlojenie, 'wb') as new_file:
                new_file.write(part.get_payload(decode=True))

# for part in msg.walk():
#     if part.get_content_disposition() == 'attachment':
#         print(part)

excel = pd.read_excel('Тест.xlsx', sheet_name='Лист1')
names = excel['Телефон '].tolist()
print(*names)
print(letter_from)
print(letter_date)


