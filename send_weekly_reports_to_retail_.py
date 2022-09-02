# !/usr/bin/python3
# -*- coding: utf-8 -*-

# Импорт системных библиотек
import os
import sys
from pathlib import Path
from datetime import datetime
# Импорт почтового сервера
import smtplib
# Импорт библиотек конструктора электронных писем
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.header import Header
# Импорт билбиотеки для чтения ini конфигов
from configparser import ConfigParser
# Ипорт библиотек подключения к mysql серверу
import pymysql
from pymysql import Error


# Функии взаимодействия с mysql
def create_connection(cfg):
    connection = None
    try:
        connection = pymysql.connect(
            host=cfg.get("mysql", "host"),
            user=cfg.get("mysql", "user"),
            passwd=cfg.get("mysql", "pass"),
            database=cfg.get("mysql", "db")
        )
        print('Connection to MySQL DB successful')
    except Error as e:
        print(f'The error {e} occurred')
    return connection


def executemany_query(connection, query, val):
    cursor = connection.cursor()
    try:
        cursor.executemany(query, val)
        connection.commit()
        res = "Query executed successfully"
        print(res)
    except Error as e:
        res = f"The error '{e}' occurred"
        print(res)
    return res


def execute_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        connection.commit()
        res = "Query executed successfully"
        print(res)
    except Error as e:
        res = f"The error '{e}' occurred"
        print(res)
    return res


def execute_read_query(connection, query):
    cursor = connection.cursor()
    try:
        cursor.execute(query)
        result = cursor.fetchall()
        return result
    except Error as e:
        print(f"The error '{e}' occurred")

# Функия отправки письма имэйл
def send_email(to_address, to_copy, shopping_centre, path_to_files_reports):
    # _получение ключей из конфига для рассылки
    str_from = cfg.get("smtp", 'from_addr')
    host = cfg.get("smtp", "server")
    pas = cfg.get("smtp", "pass")

    # _создание части письма с типом МультиПарт/Реалэйт (для основного контейнера - письмо)
    msg = MIMEMultipart('related')

    # _создание ключевых заголовков для основного контейнера (письмо)
    msg.add_header('Subject', f'Еженедельные отчеты: {shopping_centre}.')
    msg['From'] = str_from
    msg['To'] = to_address
    msg['Cc'] = to_copy
    # _создание части письма с типом МультиПарт/Альтернатив для вложения в основной контейнер (в письмо)
    msg_part_alternative = MIMEMultipart('alternative')
    msg.attach(msg_part_alternative)

    # _подготовка тела html
    html_body = """
            <html>   
              <head></head>
              <body>
                    <p>Добрый день!</p>
                    <p>Сформированы отчеты по итогам работы магазина на прошлой неделе.</p>
                    <p>При возникновении вопросов к отчетности, обращайтесь к вашему территориальному менеджеру.</p>
                    <br>
                    <br>
                    <img src="cid:{cid}"/>
                    <p><tt><b>Автоматическое формирование и рассылка отчетов</b></tt></p>               
              </body>
            </html>
        """.format(cid='image1')

    # _создание части письма для вложения в часть письма MIMEMultipart('alternative') вложенного в осн. контейнер (письмо)
    msg_part_text_html = MIMEText(html_body, 'html')
    msg_part_alternative.attach(msg_part_text_html)

    # _чтение фала в часть письма (image) для вывода в тело письма через html разметку
    path_to_logo = 'C:/Users/g.tretyachenko/PycharmProjects/ReportingService/logo.png'
    with open(path_to_logo, 'rb') as fp:
        msg_image = MIMEImage(fp.read())

    # _создание ключевого заголовка для части письма (image), вложение части в основной контейнер (письмо)
    msg_image.add_header('Content-ID', '<image1>')
    msg.attach(msg_image)
    i = 0
    for path_to_file in path_to_files_reports:
        i += 1
        # _чтение фала в часть письма (Base) для размещения во вложениии письма (добавление в основной контейнер)
        with open(path_to_file, 'rb') as fp:
            attach_file = MIMEBase('application', 'pdf', filename=os.path.basename(path_to_file))
            # _создание ключевого заголовка через класс заголовки для его декодирования с последующим указанием размещения в письме
            h = Header(os.path.basename(path_to_file), 'utf-8').encode()
            attach_file.add_header('Content-Disposition', 'attachment', filename=h)
            # _создание заголовков для индексации вложения в письмо
            attach_file.add_header('X-Attachment-Id', f'{i}')
            attach_file.add_header('Content-ID', f'<{i}>')
            # _загрузка файла в контейнер - часть письма (Base)
            attach_file.set_payload(fp.read())
            # _декодирование и добовление файла
            encoders.encode_base64(attach_file)
            msg.attach(attach_file)
    # Создание экземпляра почтового сервера и отправка письма
    server = smtplib.SMTP(host)
    server.starttls()
    server.login(msg['From'], pas)
    server.send_message(msg)
    server.quit()
    print('Отправлено!')



# Начало процедуры - подключение к ini конфигу
base_path = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(base_path, "config.ini")
if os.path.exists(config_path):
    cfg = ConfigParser()
    cfg.read(config_path)
else:
    print("Config not found! Exiting!")
    sys.exit(1)

folder_rating_goods_for_stores = Path(
    'C:/Общая/_Отчеты/!Еженедельные рассылки/Рейтинг продаж товваров за прошлую неделю (для маг.)/OutBound')
folder_anti_rating_goods_for_stores = Path(
    'C:/Общая/_Отчеты/!Еженедельные рассылки/Антирейтинг продаж товаров за прошлую неделю (для маг.)/OutBound')

# Получение списка магазинов и контактов почты для рассылки отчетов из MySQL
connection = create_connection(cfg)
sql = f'''
    SELECT 
        StoreName, 
        StoreEmail, 
        TerritorialManagerEmail, 
        MerchandisersEmail,
        OfficeEmployeesEmail,
        sub.ShoppingCentre
    FROM info_mailing_stores_contacts as info
    LEFT JOIN dim_subdivisions as sub
    ON info.SdID = sub.SdID
'''
data_frame = execute_read_query(connection, sql)
# _Готовим уникальный список почтовых адресов магазинов
email_list = set()
for row in data_frame:
    email_list.add(row[1])
email_list = [email for email in email_list]
email_list.sort()
# Готовимся и рассылаем отчеты
i = 0
dt_start_day = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0).timestamp()
for email in email_list:
    path_to_files_reports = []
    to_address = email
    to_copy = [row[2] + ', ' + row[3] + ', ' + row[4] for row in data_frame if row[1] == email]
    to_copy = set(to_copy)
    to_copy = ''.join(to_copy).replace(';', ',')
    names_stores = set(row[0] for row in data_frame if row[1] == email)
    shopping_centre = set(row[5] for row in data_frame if row[1] == email)
    report_list_1 = os.listdir(folder_rating_goods_for_stores)
    for rep in report_list_1:
        if rep.replace(' (товары-лидеры).pdf', '') in names_stores:
            if os.path.getmtime(f'{folder_rating_goods_for_stores}/{rep}') > dt_start_day:
                path_to_files_reports.append(folder_rating_goods_for_stores.joinpath(rep))
    report_list_2 = os.listdir(folder_anti_rating_goods_for_stores)
    for rep in report_list_2:
        if rep.replace(' (товары-аутсайдеры).pdf', '') in names_stores:
            if os.path.getmtime(f'{folder_anti_rating_goods_for_stores}/{rep}') > dt_start_day:
                path_to_files_reports.append(folder_anti_rating_goods_for_stores.joinpath(rep))
    if len(path_to_files_reports) > 0:
        i += 1
        print(f'----------------{i}-----------------------------')
        # Вызываем функцию отправки письма имейла
        send_email(to_address, to_copy, shopping_centre, path_to_files_reports)
        print(f'Кому: {to_address}')
        print(f'Копия: {to_copy}')
        print(f'ТЦ: {shopping_centre}')
        for att in path_to_files_reports:
            print(f'Вложение:', att)
