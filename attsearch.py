import imaplib
import quopri
import base64
from getpass import getpass

"""
Протестить.
Дописать обработку исключений.
Внедрить возможность загрузки вложения по имени.
ООП?
"""


def connect_mailbox(imap_server, email, password):
    ImapObj = imaplib.IMAP4_SSL(imap_server)
    ImapObj.login(email, password)
    ImapObj.select()
    return ImapObj


def get_emails_number(ImapObj, sender_email, attachment_format):
    """Подсчитывает количество писем, писем с вложениями, писем
    с вложениями заданного формата для указанного отправителя.
    Возвращает кортеж с перечисленными выше значениями."""

    criterion = 'FROM "{}"'.format(sender_email)
    result, data = ImapObj.search(None, criterion)
    emails_numb = len(data[0].split())

    criterion = 'X-GM-RAW "from:{} has:attachment"'.format(sender_email)
    result, data = ImapObj.search(None, criterion)
    attach_numb = len(data[0].split())

    criterion = 'X-GM-RAW "from:{} has:attachment filename:{}"'\
                .format(sender_email, attachment_format)
    result, data = ImapObj.search(None, criterion)
    spec_format_numb = len(data[0].split())
    return emails_numb, attach_numb, spec_format_numb


def get_attnames(ImapObj, email_id):
    """Находит все названия вложенных файлов, при этом не загружая
    их целиком. Возвращает объект-генератор, который на каждой итерации
    генерирует название одного вложения."""

    def decode_base64(attname_base64):
        """Предназначена для устранения ошибки incorrect padding
        при попытке декодирования. Удаляет от 1 до 4 конечных
        байт (символов)."""

        main_len = len(attname_base64[10:])
        # 10 - начало значимой части без "=?utf-8?b?"
        new_main_len = main_len - (main_len % 4 if main_len % 4 else 4)
        attname_decoded = base64.b64decode(attname_base64[10:new_main_len+10])
        return attname_decoded.decode("utf-8")

    def decode_quopri(attname_quopri):
        attname_decoded = quopri.decodestring(attname_quopri)[10:-1]
        return attname_decoded.decode("utf-8")

    def decode_attname(decode_function):
        if ' ' in attname_enc:
            attname_list = attname_enc.split()
            attname = "".join(decode_function(part)
                              for part in attname_list)
        else:
            attname = decode_function(attname_enc)
        return attname

    result, data = ImapObj.fetch(str(email_id), "(BODY)")
    email_body = data[0].decode("utf-8")
    pattern = "\"NAME\" "
    pattern_len = len(pattern)
    start = 0
    while True:
        # Прозводится поиск подстроки pattern в строке-ответе сервера.
        # Изымается подстрока в кавычках, находящаяся после подстроки pattern.
        # Дейтсвие повторяется дот тех пор, пока подстрока pattern не будет
        # отсутствовать в оставшейся части ответа.
        start = email_body.find(pattern, start)
        if start == -1:
            return
        attname_start = start + pattern_len + 1
        attname_len = email_body[attname_start:].find("\"")
        attname_end = attname_start + attname_len
        attname_enc = email_body[attname_start:attname_end]
        start += attname_len + pattern_len + 1
        if attname_enc.lower().startswith("=?utf-8?b?"):
            attname = decode_attname(decode_base64)
        elif attname_enc.lower().startswith("=?utf-8?q?"):
            attname = decode_attname(decode_quopri)
        else:
            attname = attname_enc
        yield attname


def download_attachment(ImapObj, email_id, attname, position=1):
    result, data = ImapObj.fetch(str(email_id), "(BODY[%d])" % (position+1))
    with open(attname,"wb") as f:
        f.write(base64.b64decode(data[0][1]))

    """ result, data = ImapObj.fetch(str(email_id), "(BODY)")
            email_body = data[0].decode("utf-8")
            pattern = "\"NAME\" "
            attach_numb = email_body.count(pattern)"""

if __name__ == "__main__":
    imap_server = "imap.gmail.com"
    email_login = input("Введите адрес электронной почты --- ")
    password = getpass("Введите пароль --- ")
    try:
        M = connect_mailbox(imap_server, email_login, password)
    except Exception as e:
        print("Не удалось подключиться к почтовому ящику.")
        print(e)
        print("Завершение программы...")
        exit()
    #sender = input("Введите адрес электронной почты отправителя --- ")
    #attach_format = input("Введите формат искомых файлов --- ")

    """
    emails_numb, attach_numb, spec_format_numb = \
        get_emails_number(M, sender, attach_format)
    print("Всего писем от %s -- %d" % (sender, letters_numb))
    print("Из них писем с вложениями --- %d" % attach_numb)
    print("Писем с файлом формата %s --- %d" % (attach_format, spec_format_numb))
    """
    attnames_list = []
    for name in get_attnames(M, 518):
        attnames_list.append(name.strip())
    print(518, attnames_list)
    download_attachment(M, 518, 'BHWJScqSGB4.jpg', position=2)
    M.close()
    M.logout()
