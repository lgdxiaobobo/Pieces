#!/usr/bin/python
#encoding=utf-8
import poplib

import email
from email.parser import Parser
from email.header import decode_header
from email.utils import parseaddr

import os
import time

#
# 读取QQ邮箱中的邮件
#

def guess_charset(msg):
    # get coding
    charset = msg.get_charset()
    if charset is None:
        # fail to get then get from content-type
        content_type = msg.get('Content-Type', '').lower()
        pos = content_type.find('charset=')
        if pos >= 0:
            charset = content_type[pos + 8:].strip()
    return charset

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset, 'ignore').encode('utf-8', 'ignore')
    return value

def print_info(msg):

    msg_date = msg.get('Date')
    msg_from = msg.get('From')
    msg_to   = msg.get('To')
    msg_sub  = msg.get('Subject')

    # deal with date
    date_left  = max(0, msg_date.find(',')+1)
    date_right = len(msg_date) if msg_date.find('+') == -1 else msg_date.find('+')
    date_valid = msg_date[date_left: date_right].strip(' ')
    date_tuple = time.strptime(date_valid, "%d %b %Y %H:%M:%S")
    msg_date   = str(int(time.mktime(date_tuple)))
    print msg_date
    # import pdb; pdb.set_trace()

    # deal with subject
    msg_subject  = decode_str(msg_sub)

    # deal with From and To
    _, msg_from = parseaddr(msg_from)
    _, msg_to   = parseaddr(msg_to)

    dir_name = msg_from.split('@')[0]
    if not os.path.exists(dir_name):
        os.mkdir(dir_name)

    mail_dir = os.path.join(dir_name, "%s" % (msg_date))
    if not os.path.exists(mail_dir):
        os.mkdir(mail_dir)

    conent_name = os.path.join(mail_dir, 'content')
    with open(conent_name, "w") as fp:
        fp.write("From: %s\n" % (msg_from))
        fp.write("To: %s\n" % (msg_to))
        fp.write("Date: %s\n" % (msg_date))
        fp.write("Subject: %s\n" % (msg_subject))

        j = 0
        for part in msg.walk():
            j = j + 1
            fileName = decode_str(part.get_filename())
            contentType = part.get_content_type()
            # save attachment
            if fileName != 'None':
                print fileName
                data = part.get_payload(decode=True)
                # in windows
    #            fileName = fileName.decode('utf-8').encode('gb2312', 'ignore')
                fileName = os.path.join(mail_dir, fileName)
                fEx = open(fileName, 'wb')
                fEx.write(data)
                fEx.close()
            elif contentType == 'text/plain' or contentType == 'text/html':
                fp.write('Content:\n')
                # save content
                data = part.get_payload(decode=True)
                fp.write(data)
    fp.close()

if __name__ == '__main__':

    host = 'pop.qq.com'
    user = 'XXX@qq.com'
    pwd  = 'XXXXX'

    server = poplib.POP3_SSL(host, port = '995')
    server.user(user)
    server.pass_(pwd)
    numMsgs, totalSize = server.stat()
    print "%d messages for download." % (numMsgs)

    resp, mails, octets= server.list()

    for index in range(len(mails), 0, -1):
        try:
            resp, lines, octets = server.retr(index)
            msg_content = '\r\n'.join(lines)
            msg = Parser().parsestr(msg_content)
            # print index
            print_info(msg)
        except Exception as e:
            print e

    server.quit()
