import os
import quopri
import base64
import imaplib
import zipfile
from getpass import getpass

# Произвести декомпозицию get_attnames
# Устанавливать флаги прочитанных сообщений
# Логгирование?

def connect_mailbox(email, password):
    imap_server = "imap.gmail.com"
    imap_obj = imaplib.IMAP4_SSL(imap_server)
    imap_obj.login(email, password)
    imap_obj.select()
    return imap_obj

def generate_criterion(email_sender, keywords):
    keyword_str = ' '.join(map(str, keywords))
    criterion = "from:{} has:attachment filename:{{{}}}"
    return criterion.format(email_sender, keyword_str)


def get_email_ids(imap_con, criterion):
    """Возвращает id тех писем, которые удовлетворяют критерию поиска."""

    if not email_sender:
        raise TypeError("Не указан обязательный аргумент sender_email")
    imap_con.literal = criterion.encode("utf-8")
    result, data = imap_con.search("utf-8", "X-GM-RAW")
    if result != "OK":
        raise imaplib.IMAP4_SSL.error("Не удалось произвести поиск "
                                      "по критерию %s" % criterion)
    email_ids = data[0].decode("utf-8").split()
    return [int(emid) for emid in email_ids]


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


def get_attnames(imap_con, email_id):
    """Находит все названия вложенных файлов, при этом не загружая
    их целиком. Возвращает список названий вложений для email_id."""

    def decode_attname(decode_function):
        if ' ' in attname_enc:
            attname_list = attname_enc.split()
            attname = "".join(decode_function(part)
                              for part in attname_list)
        else:
            attname = decode_function(attname_enc)
        return attname

    if not email_id:
        raise TypeError("Обязательный аргумент emid_list пуст")
    result, data = M.select()
    total_emails_numb = int(data[0].decode("utf-8"))
    if email_id > total_emails_numb:
        raise imaplib.IMAP4_SSL.error("Указан некорректный email_id", email_id)
    result, data = imap_con.fetch(str(email_id), "(BODY)")
    email_body = data[0].decode("utf-8")
    pattern = "\"NAME\" "
    pattern_len = len(pattern)
    start = 0
    emid_attnames = []
    while True:
        # Прозводится поиск подстроки pattern в строке-ответе сервера.
        # Изымается подстрока в кавычках, находящаяся после подстроки pattern.
        # Дейтсвие повторяется дот тех пор, пока подстрока pattern не будет
        # отсутствовать в оставшейся части ответа.
        start = email_body.find(pattern, start)
        if start == -1:
            break
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
        emid_attnames.append(attname)
    return emid_attnames


def download_attachment(imap_con, email_id, attname_list):
    dirname = "PRICES"
    current_path = os.path.join(os.environ['HOME'], dirname)
    if not os.path.exists(current_path):
        os.makedirs(current_path)
    for pos, attname in enumerate(attname_list, 1):
        result, data = imap_con.fetch(str(email_id), "(BODY[%d])" % (pos+1))
        filepath = os.path.join(current_path, attname)
        with open(filepath, "wb") as f:
            f.write(base64.b64decode(data[0][1]))
        """filename, file_extension = os.path.splitext(filepath)
        if file_extension == ".zip":
            with zipfile.ZipFile(filepath, "r") as myzip:
                myzip.extractall(current_path)
            os.remove(filepath)"""


def close_connection(imap_con):
    imap_con.close()
    imap_con.logout()

if __name__ == "__main__":
    email_sender = "kacaruba.yura@mail.ru"
    keywords = ['альфа', 'list_opt', 'эксмо', 'аст']
    ids_criterion = generate_criterion(email_sender, keywords)
    email_login = input("Введите адрес электронной почты --- ")
    password = getpass("Введите пароль --- ")
    try:
        M = connect_mailbox(email_login, password)
    except (imaplib.IMAP4_SSL.error, TypeError) as error:
        print("Не удалось подключиться к почтовому ящику.")
        print(error)
        exit("Завершение программы...")
    try:
        email_ids = get_email_ids(M, ids_criterion)
        for emid in email_ids:
            attname_list = get_attnames(M, emid)
            # download_attachment(M, emid, attname_list)
    except (imaplib.IMAP4_SSL.error, TypeError) as error:
        print(error)
        exit("Завершение программы...")
    close_connection(M)
