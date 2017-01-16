import imaplib
from getpass import getpass


def get_letters_number(ImapObj, sender_email, attachment_format):
    criterion = 'FROM "{}"'.format(sender_email)
    result, data = ImapObj.search(None, criterion)
    letters_numb = len(data[0].split())

    criterion = 'X-GM-RAW "from:{} has:attachment"'.format(sender_email)
    result, data = ImapObj.search(None, criterion)
    attach_numb = len(data[0].split())

    criterion = 'X-GM-RAW "from:{} has:attachment filename:{}"'\
                .format(sender_email,attachment_format)
    result, data = ImapObj.search(None, criterion)
    spec_format_numb = len(data[0].split())
    return letters_numb, attach_numb, spec_format_numb

if __name__ == "__main__":
    imap_server = "imap.gmail.com"
    email_login = input("Введите адрес электронной почты(gmail) --- ")
    password = getpass("Введите пароль --- ")
    sender = input("Введите адрес электронной почты отправителя --- ")
    attach_format = input("Введите формат искомых файлов --- ")
    M = imaplib.IMAP4_SSL(imap_server)
    M.login(email_login, password)
    M.select()
    letters_numb, attach_numb, spec_format_numb = \
        get_letters_number(M, sender, attach_format)
    print("Всего писем от %s -- %d" % (sender, letters_numb))
    print("Из них писем с вложениями --- %d" % attach_numb)
    print("Писем с файлом формата %s --- %d" % (attach_format, spec_format_numb))
    M.close()
    M.logout()
    
