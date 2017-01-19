import imaplib
import quopri
import base64
from getpass import getpass


"""
Всё тщательно проверить.
Протестить.
Дописать обработку исключений.
Внедрить возможность загрузки вложения по имени.
"""

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

def get_attachs_name(ImapObj, letter_id):
    def decode_base64(attachname_base64):
        lens = len(attachname_base64[10:])
        lenx = lens - (lens % 4 if lens % 4 else 4)
        attachname_decoded = base64.b64decode(attachname_base64[10:lenx+10])
        return attachname_decoded.decode("utf-8")
    def decode_quopri(attachname_quopri):
        attachname_decoded = quopri.decodestring(attachname_quopri)[10:-1]
        return attachname_decoded.decode("utf-8")
    result, data = M.fetch(str(letter_id), "(BODY)")
    letter_body = data[0].decode()
    pattern = "\"NAME\" "
    pattern_l = len(pattern)
    start = 0
    while True:
        start = letter_body.find(pattern, start)
        if start == -1:
            return
        attachname_start = start + pattern_l + 1
        attachname_end = attachname_start + letter_body[attachname_start:].find("\"")
        attachname_enc = letter_body[attachname_start:attachname_end]
        start += letter_body[attachname_start:].find("\"") + pattern_l + 1
        if attachname_enc.lower().startswith("=?utf-8?b?"):    #TO FIND SHORTER SOLUTION
            if ' ' in attachname_enc:
                attachname_list = attachname_enc.split()
                attachname = "".join(decode_base64(part)
                                     for part in attachname_list)
            else:
                attachname = decode_base64(attachname_enc)
        elif attachname_enc.lower().startswith("=?utf-8?q?"):
            if ' ' in attachname_enc:
                attachname_list = attachname_enc.split()
                attachname = "".join(decode_quopri(part)
                                     for part in attachname_list)
            else:
                attachname = decode_quopri(attachname_enc)
        else:
            attachname = attachname_enc
        yield attachname


if __name__ == "__main__":
    imap_server = "imap.gmail.com"
    email_login = input("Введите адрес электронной почты --- ")
    password = getpass("Введите пароль --- ")
    #sender = input("Введите адрес электронной почты отправителя --- ")
    #attach_format = input("Введите формат искомых файлов --- ")
    M = imaplib.IMAP4_SSL(imap_server)
    M.login(email_login, password)
    M.select()
    """
    letters_numb, attach_numb, spec_format_numb = \
        get_letters_number(M, sender, attach_format)
    print("Всего писем от %s -- %d" % (sender, letters_numb))
    print("Из них писем с вложениями --- %d" % attach_numb)
    print("Писем с файлом формата %s --- %d" % (attach_format, spec_format_numb))
    """
    for name in get_attachs_name(M, 515):
        print(name)
    M.close()
    M.logout()
    
