import imaplib
import email
from datetime import datetime, timedelta
import os
import fnmatch

date = (datetime.today()).strftime("%d-%b-%Y")
today = datetime.today()
cutoff = today - timedelta(days=7)
dt = cutoff.strftime('%d-%b-%Y')
error_file = []

'''
f_name = ["DESB", "Conoco Phillips"]
m_word = ["*DESB Vessel Summary Daily Report*", "*ConocoPhillips Vessel Summary Daily*"]
m_file = ["DESB Vessel Summary Daily*", "*ConocoPhillips Vessel Summary Daily*"]

f_name = ["Hess", "Petrofac", "Jxnippon", "PFLNG Satu", "PTTEP"]
m_word = ["*HESS Malaysia - Vessel Summary Daily Report*", "*PetrofacMalaysia Daily Summary*", "*JXNippon Vessel Summary*", "PFLNG Satu Summary Daily Report*", "All PTTEP Summary*"]
m_file = ["*HESS Malaysia -  Vessel Summary*", "*Vessel Summary Daily Report*", "*JXNippon Vessel Summary*", "PFLNG Satu - IDP Vessel Summary Daily Report*", "*Vessel Summary Daily Report*"]
'''
f_name = ["DAILY SUMMARY"]
m_word = ["PetrofacMalaysia Daily Summary*"]
m_file = ["*PetrofacMalaysia - Vessel Summary Daily Report*"]

file_path = r'C:\Users\mvmwe\Dropbox\MVMCC\REPORT\WEEKLY\PETROFAC'

email_add = "mvmcc@wild-geese-group.com"
password = "s9nPviD\\"

if not os.path.exists(file_path):
    os.mkdir(file_path)

try:
    imap = imaplib.IMAP4_SSL("mail.wild-geese-group.com", 993)
    imap.login(email_add, password)
    imap.select('Inbox')
    #typ, data = imap.search(None, '(SENTSINCE {0})'.format(date))

    index = 0
    while index < len(f_name) and index < len(m_word):
        folder_name = os.path.join(file_path, f_name[index])
        if not os.path.exists(folder_name):
            os.mkdir(folder_name)

        typ, data = imap.search(None, '(SINCE %s)' % (dt,), '(FROM mvmcc@wild-geese-group.com)')

        for num in data[0].split():
            typ, data = imap.fetch(num, '(RFC822)')
            raw_email = data[0][1]
            raw_email_string = raw_email.decode('windows-1252')
            email_message = email.message_from_string(raw_email_string)
            subject_name = email_message['subject']

            att_path = "No attachment found from email " + subject_name
            for part in email_message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                fileName = part.get_filename()

                if fnmatch.fnmatch(subject_name, m_word[index]):
                    if fnmatch.fnmatch(fileName, m_file[index]):
                        att_path = os.path.join(folder_name, fileName)
                        print(att_path)
                        if not os.path.isfile(att_path):
                            fp = open(att_path, "wb")
                            fp.write(part.get_payload(decode=True))
                            fp.close()

            print(att_path)
        index += 1

    imap.close()
    imap.logout()
except OSError as e:
    print(e)
    error_file.append(e)
    pass
